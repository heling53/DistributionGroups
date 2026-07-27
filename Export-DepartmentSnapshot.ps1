#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Снимок состояния подразделений 1С и групп рассылки AD. Ничего не изменяет.
.DESCRIPTION
    Пишет в $SnapshotPath файл DeptSnapshot_<дата>.json со срезом:
      - подразделения из 1С (GUID, полное наименование, вершина, руководитель);
      - пользователи AD (табельный номер -> подразделение);
      - группы рассылки (имя, SID, SamAccountName, mail, состав).

    Зачем: 1С не переименовывает подразделения, а заводит новое и удаляет старое,
    поэтому сопоставить старую группу с новой по идентификатору 1С невозможно.
    Единственный признак, который переживает такую «замену», — состав людей
    (табельные номера). Чтобы его сравнивать, нужна история: снимок ДО и ПОСЛЕ.
    Скрипт накапливает эту историю. Сам по себе он ничего не решает и ничего
    не меняет — это регистратор для последующего сопоставления и для
    восстановления ошибочно удалённой группы (SID + состав сохранены).

    Запускать по расписанию ПЕРЕД Sync-DistributionGroups.ps1.
.EXAMPLE
    .\Export-DepartmentSnapshot.ps1 -Verbose
#>

[CmdletBinding()]
param(
    [string]$ExcelPath          = 'D:\Scripts\verify1C\1c\Подразделения.xlsx',
    [string]$ExcelSheet         = 'Лист_1',
    [string]$DcServer           = 'srv-dc3.transitcard.ru',
    [string]$TargetOU           = 'OU=Группы рассылки,DC=transitcard,DC=ru',
    [string]$GroupDesc          = 'Группа рассылки подразделения_v1.0',
    [string[]]$UserSearchBases  = @('OU=Users.Corp,DC=transitcard,DC=ru',
                                    'OU=Users.Contractor,DC=transitcard,DC=ru'),
    [string]$SnapshotPath       = 'D:\Scripts\Группы рассылки\Snapshots'
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) { 'ERROR' { 'Red' }; 'WARN' { 'Yellow' }; 'OK' { 'Green' }; default { 'Gray' } }
    Write-Host ("[{0:HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message) -ForegroundColor $color
}

# Значение колонки по точному имени, с запасным поиском по маске:
# заголовки в выгрузке 1С содержат опечатки («пдоразделения») и могут меняться.
function Get-Column {
    param($Row, [string]$Exact, [string]$Like)
    $v = $Row.$Exact
    if (-not $v -and $Like) {
        $v = $Row.PSObject.Properties |
             Where-Object { $_.Name -like $Like } |
             Select-Object -First 1 -ExpandProperty Value
    }
    if ($v) { ($v.ToString() -replace '\s+', ' ').Trim() } else { '' }
}

Import-Module ActiveDirectory
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "Модуль ImportExcel не установлен. Установите: Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel

if (-not (Test-Path $SnapshotPath)) { New-Item -ItemType Directory -Path $SnapshotPath -Force | Out-Null }

# ==========================================
# 1С: справочник подразделений
# ==========================================
Write-Log "Чтение справочника 1С: $ExcelPath"
if (-not (Test-Path $ExcelPath)) { throw "Справочник подразделений не найден: $ExcelPath" }

$departments = New-Object System.Collections.Generic.List[object]
foreach ($row in (Import-Excel -Path $ExcelPath -WorksheetName $ExcelSheet)) {
    $fullName = Get-Column -Row $row -Exact 'Полное наименование пдоразделения' -Like '*полное*'
    if ([string]::IsNullOrWhiteSpace($fullName)) { continue }
    # Строка 2 выгрузки — технические имена колонок (objFullName/objName), не данные.
    if ($fullName -match 'objName|objFullName|^obj|полное наименование') { continue }

    $managerEN = Get-Column -Row $row -Exact 'Руководитель подразделения' -Like '*руководитель*'
    # Ведущие нули Excel-формата ("0000003966") срезаются, кроме табельного
    # "0000-00013": в AD он хранится с нулями и дефисом.
    if ($managerEN -and $managerEN -ne '0000-00013') { $managerEN = $managerEN -replace '^0+', '' }

    $departments.Add([PSCustomObject]@{
        Id            = Get-Column -Row $row -Exact 'Уникальный идентификатор подразделения' -Like '*идентификатор*'
        Name          = $fullName
        Parent        = Get-Column -Row $row -Exact 'Наименование подразделения (вершины)' -Like '*вершины*'
        ManagerEmpNum = $managerEN
    })
}
Write-Log "1С: подразделений $($departments.Count)" 'OK'

# ==========================================
# AD: пользователи
# ==========================================
# Табельный номер (EmployeeNumber) — стабильный ключ человека: он переживает
# и переименование подразделения, и смену ФИО, и смену SamAccountName.
Write-Log "Чтение пользователей AD ($DcServer)"
$users     = New-Object System.Collections.Generic.List[object]
$samByDN   = @{}
foreach ($ou in $UserSearchBases) {
    $adUsers = Get-ADUser -SearchBase $ou -Server $DcServer `
                          -Filter "Enabled -eq 'True' -and Department -like '*'" `
                          -Properties EmployeeNumber, Department, Name, SamAccountName `
                          -ResultPageSize 1000
    foreach ($u in $adUsers) {
        # Те же исключения, что и в синхронизации: служебные и региональные учётки.
        if ($u.Name -match '\[ADM\]|\[VPN\]|\[KZ\]')                 { continue }
        if ($u.SamAccountName -match '^(adm|vpn)-|test')             { continue }
        if ($u.Department -notmatch '\S')                            { continue }

        $samByDN[$u.DistinguishedName] = $u.SamAccountName
        $users.Add([PSCustomObject]@{
            EmpNum     = if ($u.EmployeeNumber) { $u.EmployeeNumber.ToString().Trim() } else { '' }
            Sam        = $u.SamAccountName
            Name       = $u.Name
            Department = ($u.Department -replace '\s+', ' ').Trim()
        })
    }
}
Write-Log "AD: пользователей $($users.Count)" 'OK'

# ==========================================
# AD: группы рассылки
# ==========================================
# SID и состав пишем намеренно: если группа будет удалена ошибочно, снимок —
# единственный источник, по которому её можно восстановить один в один
# (SID нужен, чтобы понять, какие ACL и группы доступа осиротели).
Write-Log "Чтение групп рассылки ($TargetOU)"
$adGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" -SearchBase $TargetOU -Server $DcServer `
                        -Properties Name, DisplayName, SamAccountName, mail, Member, ManagedBy, whenCreated, objectSid

$groupNameByDN = @{}
foreach ($g in $adGroups) { $groupNameByDN[$g.DistinguishedName] = $g.Name }

$groups = New-Object System.Collections.Generic.List[object]
foreach ($g in $adGroups) {
    $memberUsers  = New-Object System.Collections.Generic.List[string]
    $memberGroups = New-Object System.Collections.Generic.List[string]
    foreach ($dn in $g.Member) {
        if     ($groupNameByDN.ContainsKey($dn)) { $memberGroups.Add($groupNameByDN[$dn]) }
        elseif ($samByDN.ContainsKey($dn))       { $memberUsers.Add($samByDN[$dn]) }
        else                                     { $memberUsers.Add($dn) }   # вне снимка: отключён/иная OU
    }
    $groups.Add([PSCustomObject]@{
        Name         = $g.Name
        DisplayName  = $g.DisplayName
        Sam          = $g.SamAccountName
        Sid          = if ($g.objectSid) { $g.objectSid.Value } else { '' }
        Mail         = $g.mail
        ManagedBy    = $g.ManagedBy
        WhenCreated  = if ($g.whenCreated) { $g.whenCreated.ToString('o') } else { '' }
        MemberUsers  = $memberUsers.ToArray()
        MemberGroups = $memberGroups.ToArray()
    })
}
Write-Log "AD: групп рассылки $($groups.Count)" 'OK'

# ==========================================
# Запись снимка
# ==========================================
$snapshot = [PSCustomObject]@{
    Timestamp   = (Get-Date).ToString('o')
    Source      = [PSCustomObject]@{ Excel = $ExcelPath; DcServer = $DcServer; TargetOU = $TargetOU }
    Departments = $departments.ToArray()
    Users       = $users.ToArray()
    Groups      = $groups.ToArray()
}

$outFile = Join-Path $SnapshotPath ("DeptSnapshot_{0:yyyyMMdd_HHmmss}.json" -f (Get-Date))
$snapshot | ConvertTo-Json -Depth 6 | Set-Content -Path $outFile -Encoding UTF8
Write-Log "Снимок записан: $outFile" 'OK'
