<#
.SYNOPSIS
    Прогон настоящего Sync-DistributionGroups.ps1 против каталога Active Directory в памяти.
.DESCRIPTION
    Домен и Exchange не нужны: модули ActiveDirectory и ImportExcel в tests/modules —
    поддельные, они держат каталог в $global:ADStore, а команды Exchange заглушены здесь.
    Проверяется настоящий скрипт целиком, без правок под тесты.
.EXAMPLE
    pwsh -NoProfile -File .\tests\Run-Tests.ps1
#>
$ErrorActionPreference = 'Stop'
$SP     = $PSScriptRoot
$Script = Join-Path (Split-Path $PSScriptRoot -Parent) 'Sync-DistributionGroups.ps1'
$Work   = Join-Path $SP 'work'
$env:PSModulePath = (Join-Path $SP 'modules') + [IO.Path]::PathSeparator + $env:PSModulePath

$OU      = 'OU=Группы рассылки,DC=transitcard,DC=ru'
$CorpOU  = 'OU=Users.Corp,DC=transitcard,DC=ru'
$Desc    = 'Группа рассылки подразделения_v1.0'

# ---------- Заглушки Exchange и транскрипта ----------
function New-PSSession    { param($ConfigurationName,$ConnectionUri,$Authentication) return [PSCustomObject]@{ Id = 1 } }
function Import-PSSession  { param($Session,[switch]$DisableNameChecking,[switch]$AllowClobber) }
function Remove-PSSession  { [CmdletBinding()] param([Parameter(ValueFromPipeline)]$Session) }
function Start-Transcript  { param($Path,[switch]$Append) }
function Stop-Transcript   { }
function Enable-DistributionGroup {
    param($Identity,$Alias,$PrimarySmtpAddress,$DomainController)
    $g = $global:ADStore.Groups | Where-Object { $_.SamAccountName -eq $Identity -or $_.Name -eq $Identity }
    if (-not $g) { throw "Enable-DistributionGroup: не найдена $Identity" }
    $g.mail = $PrimarySmtpAddress
}
function Set-DistributionGroup {
    param($Identity,$RequireSenderAuthenticationEnabled,$HiddenFromAddressListsEnabled,$DomainController,$DisplayName,$Name)
    $g = $global:ADStore.Groups | Where-Object { $_.SamAccountName -eq $Identity -or $_.Name -eq $Identity -or $_.DistinguishedName -eq $Identity }
    if (-not $g) { throw "Set-DistributionGroup: не найдена $Identity" }
    if ($null -ne $HiddenFromAddressListsEnabled) { $g.HiddenFromAddressListsEnabled = $HiddenFromAddressListsEnabled }
}
function Disable-DistributionGroup {
    param($Identity,[switch]$Confirm,$DomainController)
    $g = $global:ADStore.Groups | Where-Object { $_.DistinguishedName -eq $Identity -or $_.SamAccountName -eq $Identity }
    if ($g) { $g.mail = $null }
}

# ---------- Инфраструктура тестов ----------
$script:Pass = 0; $script:Fail = 0; $script:Failures = @()
function Check {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { $script:Pass++; "   [ok]   $Name" }
    else { $script:Fail++; $script:Failures += $Name; "   [FAIL] $Name $Detail" }
}

function Reset-Store {
    param([datetime]$Now = (Get-Date))
    $global:ADStore = @{ Users = @(); Groups = @(); NextSid = 100; Now = $Now }
    $global:ExcelRows = @()
    if (Test-Path $Work) { Remove-Item $Work -Recurse -Force }
    New-Item -ItemType Directory -Path $Work -Force | Out-Null
    Set-Content -Path (Join-Path $Work 'Подразделения.xlsx') -Value 'stub' -Encoding UTF8
    Set-Content -Path (Join-Path $Work 'protected_groups.json') -Value '[]' -Encoding UTF8
}

function Add-User {
    param($Sam, $Name, $Dept, $EmpNum, $Ou = $CorpOU, [bool]$Enabled = $true)
    $global:ADStore.Users += [PSCustomObject]@{
        DistinguishedName = "CN=$Name,$Ou"
        SamAccountName    = $Sam
        Name              = $Name
        Department        = $Dept
        EmployeeNumber    = $EmpNum
        Enabled           = $Enabled
    }
}

function Set-Excel {
    param($Rows)   # @( @{ Child='...'; Parent='...'; Manager='...' } )
    $global:ExcelRows = @($Rows | ForEach-Object {
        [PSCustomObject]@{
            objFullName                  = $_.Child
            objName                      = $_.Parent
            'Руководитель подразделения' = $_.Manager
        }
    })
}

function Invoke-Sync {
    param([switch]$ShowOutput)
    $out = & $Script `
        -ExcelPath               (Join-Path $Work 'Подразделения.xlsx') `
        -LogPath                 (Join-Path $Work 'Logs') `
        -ProtectedGroupsJsonPath (Join-Path $Work 'protected_groups.json') `
        -AliasesJsonPath         (Join-Path $Work 'department_aliases.json') `
        -RegistryJsonPath        (Join-Path $Work 'department_registry.json') `
        -HistoryJsonPath         (Join-Path $Work 'department_history.json') `
        -RenameReportCsv         (Join-Path $Work 'RenameCandidates.csv') `
        -TargetOU $OU -UsersCorpOU $CorpOU -GroupDesc $Desc *>&1 | Out-String
    if ($ShowOutput) { $out }
    $script:LastOutput = $out
}

function Get-G      { param($Name) $global:ADStore.Groups | Where-Object { $_.Name -eq $Name } | Select-Object -First 1 }
function Get-Mem    { param($Name) $g = Get-G $Name; if ($g) { @($g.Member) } else { @() } }
function Get-UserDN { param($Sam) ($global:ADStore.Users | Where-Object { $_.SamAccountName -eq $Sam }).DistinguishedName }
function Get-MemSams {
    param($Name)
    $dns = Get-Mem $Name
    @($global:ADStore.Users | Where-Object { $dns -contains $_.DistinguishedName } | ForEach-Object { $_.SamAccountName }) | Sort-Object
}
function Read-Json { param($File) $p = Join-Path $Work $File; if (Test-Path $p) { Get-Content $p -Raw | ConvertFrom-Json } else { $null } }

# Базовая организация: коммерческий блок с двумя отделами.
function Build-BaseOrg {
    param([int]$SalesCount = 12, [int]$PurchCount = 5)
    Set-Excel @(
        @{ Child = 'Отдел продаж';  Parent = 'Коммерческий блок'; Manager = '1001' },
        @{ Child = 'Отдел закупок'; Parent = 'Коммерческий блок'; Manager = '2001' }
    )
    Add-User -Sam 'boss.sales' -Name 'Руководитель Продаж' -Dept 'Отдел продаж'  -EmpNum '1001'
    Add-User -Sam 'boss.purch' -Name 'Руководитель Закупок' -Dept 'Отдел закупок' -EmpNum '2001'
    for ($i = 1; $i -lt $SalesCount; $i++) { Add-User -Sam "sales$i" -Name "Продавец $i" -Dept 'Отдел продаж'  -EmpNum "10$i" }
    for ($i = 1; $i -lt $PurchCount; $i++) { Add-User -Sam "purch$i" -Name "Закупщик $i" -Dept 'Отдел закупок' -EmpNum "20$i" }
}

function Move-Users {
    param([string[]]$Sams, [string]$ToDept)
    foreach ($s in $Sams) {
        ($global:ADStore.Users | Where-Object { $_.SamAccountName -eq $s }).Department = $ToDept
    }
}

'============================================================'
'ТЕСТ 1. Первичное создание: группы, состав, иерархия'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
Check 'создана группа «Отдел продаж»'   ((Get-G 'Отдел продаж')  -ne $null)
Check 'создана группа «Отдел закупок»'  ((Get-G 'Отдел закупок') -ne $null)
Check 'создана вершина «Коммерческий блок»' ((Get-G 'Коммерческий блок') -ne $null)
Check 'в продажах 12 человек' ((Get-MemSams 'Отдел продаж').Count -eq 12) "= $((Get-MemSams 'Отдел продаж').Count)"
Check 'в закупках 5 человек'  ((Get-MemSams 'Отдел закупок').Count -eq 5) "= $((Get-MemSams 'Отдел закупок').Count)"
$blok = Get-G 'Коммерческий блок'
Check 'вершина содержит обе дочерние группы' (@($blok.Member).Count -eq 2) "= $(@($blok.Member).Count)"
Check 'ManagedBy проставлен' ((Get-G 'Отдел продаж').ManagedBy -eq (Get-UserDN 'boss.sales'))
Check 'адрес выдан' ((Get-G 'Отдел продаж').mail -like '*@transitcard.ru')
Check 'реестр создан на 3 группы' (@(Read-Json 'department_registry.json').Count -eq 3)

''
'============================================================'
'ТЕСТ 2. Повторный прогон не меняет ничего (идемпотентность)'
'============================================================'
$before = ($global:ADStore.Groups | ForEach-Object { "$($_.Name)|$($_.SamAccountName)|$(@($_.Member) -join ',')" }) -join ';'
Invoke-Sync
$after  = ($global:ADStore.Groups | ForEach-Object { "$($_.Name)|$($_.SamAccountName)|$(@($_.Member) -join ',')" }) -join ';'
Check 'состояние каталога не изменилось' ($before -eq $after)
Check 'в журнале нет новых событий' (-not ($script:LastOutput -match 'Переименована|Создана группа|Удалена'))

''
'============================================================'
'ТЕСТ 3. Полное переименование за один день'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$sidBefore  = (Get-G 'Отдел продаж').objectSid.Value
$samBefore  = (Get-G 'Отдел продаж').SamAccountName
$mailBefore = (Get-G 'Отдел продаж').mail
# 1С: отдел заведён заново под новым названием, старый удалён.
Set-Excel @(
    @{ Child = 'Отдел продаж и маркетинга'; Parent = 'Коммерческий блок'; Manager = '1001' },
    @{ Child = 'Отдел закупок';             Parent = 'Коммерческий блок'; Manager = '2001' }
)
Move-Users -Sams (@('boss.sales') + (1..11 | ForEach-Object { "sales$_" })) -ToDept 'Отдел продаж и маркетинга'
Invoke-Sync
$new = Get-G 'Отдел продаж и маркетинга'
Check 'группа переименована, а не создана заново' ($new -ne $null -and $new.objectSid.Value -eq $sidBefore)
Check 'SamAccountName сохранён'  ($new.SamAccountName -eq $samBefore)
Check 'почтовый адрес сохранён'  ($new.mail -eq $mailBefore)
Check 'старой группы больше нет' ((Get-G 'Отдел продаж') -eq $null)
Check 'дубль не создан' (@($global:ADStore.Groups | Where-Object { $_.Description -eq $Desc }).Count -eq 3)
Check 'состав на месте (12)' ((Get-MemSams 'Отдел продаж и маркетинга').Count -eq 12) "= $((Get-MemSams 'Отдел продаж и маркетинга').Count)"
Check 'вложенность в вершину сохранена' (@((Get-G 'Коммерческий блок').Member) -contains $new.DistinguishedName)
$hist = Read-Json 'department_history.json'
Check 'событие Rename записано в журнал' (@($hist | Where-Object { $_.Event -eq 'Rename' }).Count -eq 1)
$reg = Read-Json 'department_registry.json'
$entry = $reg | Where-Object { $_.Sam -eq $samBefore }
Check 'в реестре два названия с историей' (@($entry.Names).Count -eq 2) "= $(@($entry.Names).Count)"

''
'============================================================'
'ТЕСТ 4. Постепенный переезд: 4 из 12 в первый день'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$sidBefore = (Get-G 'Отдел продаж').objectSid.Value
Set-Excel @(
    @{ Child = 'Отдел продаж и маркетинга'; Parent = 'Коммерческий блок'; Manager = '1001' },
    @{ Child = 'Отдел закупок';             Parent = 'Коммерческий блок'; Manager = '2001' }
)
Move-Users -Sams @('sales1','sales2','sales3','sales4') -ToDept 'Отдел продаж и маркетинга'
Invoke-Sync
$new = Get-G 'Отдел продаж и маркетинга'
Check 'переименование распознано по 4 ушедшим' ($new -ne $null -and $new.objectSid.Value -eq $sidBefore)
Check 'две группы одновременно НЕ существуют' ((Get-G 'Отдел продаж') -eq $null)
Check 'все 12 в группе: 4 переехавших + 8 по псевдониму' ((Get-MemSams 'Отдел продаж и маркетинга').Count -eq 12) "= $((Get-MemSams 'Отдел продаж и маркетинга').Count)"
$al = Read-Json 'department_aliases.json'
Check 'псевдоним записан' (@($al | Where-Object { $_.Old -eq 'Отдел продаж' -and $_.New -eq 'Отдел продаж и маркетинга' }).Count -eq 1)

# Дни 2 и 3: доезжают остальные — состав всё это время полный.
Move-Users -Sams @('sales5','sales6','sales7','sales8') -ToDept 'Отдел продаж и маркетинга'
Invoke-Sync
Check 'день 2: по-прежнему все 12' ((Get-MemSams 'Отдел продаж и маркетинга').Count -eq 12) "= $((Get-MemSams 'Отдел продаж и маркетинга').Count)"
Move-Users -Sams @('sales9','sales10','sales11','boss.sales') -ToDept 'Отдел продаж и маркетинга'
Invoke-Sync
Check 'день 3: все переехали, состав 12' ((Get-MemSams 'Отдел продаж и маркетинга').Count -eq 12) "= $((Get-MemSams 'Отдел продаж и маркетинга').Count)"
Check 'групп по-прежнему 3' (@($global:ADStore.Groups | Where-Object { $_.Description -eq $Desc }).Count -eq 3)

''
'============================================================'
'ТЕСТ 5. Разделение отдела надвое — автоматом НЕ переименовывать'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$sidBefore = (Get-G 'Отдел продаж').objectSid.Value
Set-Excel @(
    @{ Child = 'Отдел продаж СНГ';    Parent = 'Коммерческий блок'; Manager = '1001' },
    @{ Child = 'Отдел продаж Европа'; Parent = 'Коммерческий блок'; Manager = '1005' },
    @{ Child = 'Отдел закупок';       Parent = 'Коммерческий блок'; Manager = '2001' }
)
Move-Users -Sams @('boss.sales','sales1','sales2','sales3','sales4','sales5') -ToDept 'Отдел продаж СНГ'
Move-Users -Sams @('sales6','sales7','sales8','sales9','sales10','sales11') -ToDept 'Отдел продаж Европа'
Invoke-Sync
Check 'старая группа не переименована автоматически' ((Get-G 'Отдел продаж') -ne $null)
Check 'создана группа СНГ'    ((Get-G 'Отдел продаж СНГ') -ne $null)
Check 'создана группа Европа' ((Get-G 'Отдел продаж Европа') -ne $null)
Check 'состав старой группы заморожен (12)' ((Get-MemSams 'Отдел продаж').Count -eq 12) "= $((Get-MemSams 'Отдел продаж').Count)"
Check 'старая группа в карантине' ((Get-G 'Отдел продаж').info -match 'QUARANTINE:')
Check 'старая группа скрыта из адресной книги' ((Get-G 'Отдел продаж').HiddenFromAddressListsEnabled -eq $true)
$csv = Import-Csv (Join-Path $Work 'RenameCandidates.csv')
Check 'спорная пара помечена Manual' (@($csv | Where-Object { $_.OldGroup -eq 'Отдел продаж' -and $_.Decision -eq 'Manual' }).Count -eq 1) "решения: $(($csv | ForEach-Object { $_.Decision }) -join ',')"

''
'============================================================'
'ТЕСТ 6. Обычный перевод сотрудника — не переименование'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
Move-Users -Sams @('sales1') -ToDept 'Отдел закупок'
Invoke-Sync
Check 'ничего не переименовано' ((Get-G 'Отдел продаж') -ne $null -and (Get-G 'Отдел закупок') -ne $null)
Check 'сотрудник ушёл из продаж' (-not ((Get-MemSams 'Отдел продаж') -contains 'sales1'))
Check 'сотрудник добавлен в закупки' ((Get-MemSams 'Отдел закупок') -contains 'sales1')
Check 'в продажах 11' ((Get-MemSams 'Отдел продаж').Count -eq 11) "= $((Get-MemSams 'Отдел продаж').Count)"
Check 'в закупках 6'  ((Get-MemSams 'Отдел закупок').Count -eq 6) "= $((Get-MemSams 'Отдел закупок').Count)"

''
'============================================================'
'ТЕСТ 7. Роспуск отдела: карантин, затем удаление по истечении'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$sidPurch = (Get-G 'Отдел закупок').objectSid.Value
Set-Excel @( @{ Child = 'Отдел закупок'; Parent = 'Коммерческий блок'; Manager = '2001' } )
Move-Users -Sams (@('boss.sales') + (1..11 | ForEach-Object { "sales$_" })) -ToDept 'Отдел закупок'
Invoke-Sync
Check 'группа не удалена сразу' ((Get-G 'Отдел продаж') -ne $null)
Check 'живой «Отдел закупок» не захвачен распавшимся отделом' ((Get-G 'Отдел закупок').objectSid.Value -eq $sidPurch)
Check 'никого не отставили в ZZ-Дубль' (@($global:ADStore.Groups | Where-Object { $_.Name -like 'ZZ-*' }).Count -eq 0)
Check 'группа в карантине' ((Get-G 'Отдел продаж').info -match 'QUARANTINE:')
Check 'состав заморожен, доступ жив' ((Get-MemSams 'Отдел продаж').Count -eq 12) "= $((Get-MemSams 'Отдел продаж').Count)"
# Отматываем дату карантина на 40 дней назад — как будто прошёл месяц.
(Get-G 'Отдел продаж').info = "QUARANTINE:{0:yyyy-MM-dd}" -f (Get-Date).AddDays(-40)
Invoke-Sync
Check 'по истечении карантина группа удалена' ((Get-G 'Отдел продаж') -eq $null)
Check 'после удаления «Отдел закупок» на месте' ((Get-G 'Отдел закупок').objectSid.Value -eq $sidPurch)
Check 'ZZ-Дубль так и не появился' (@($global:ADStore.Groups | Where-Object { $_.Name -like 'ZZ-*' }).Count -eq 0)
$hist = Read-Json 'department_history.json'
$del = @($hist | Where-Object { $_.Event -eq 'Delete' })
Check 'в журнале есть Delete с SID' ($del.Count -eq 1 -and $del[0].Sid)
Check 'в журнале сохранён состав удалённой группы' (@($del[0].Members).Count -eq 12) "= $(@($del[0].Members).Count)"

''
'============================================================'
'ТЕСТ 8. Дубль из прошлого: перенос вложенности в группу Jira'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$sidOld = (Get-G 'Отдел продаж').objectSid.Value
# Имитируем вчерашний прогон старой версии: дубль уже создан, люди частично переехали.
$global:ADStore.Groups += [PSCustomObject]@{
    DistinguishedName = "CN=Отдел продаж и маркетинга,$OU"
    Name = 'Отдел продаж и маркетинга'; SamAccountName = 'DG_OPIM'
    DisplayName = 'Отдел продаж и маркетинга'; Description = $Desc
    Member = @(); MemberOf = @(); ManagedBy = $null
    mail = 'dg_opim@transitcard.ru'; info = $null
    whenCreated = (Get-Date).AddDays(-3)
    objectSid = [PSCustomObject]@{ Value = 'S-1-5-21-1-1-1-999' }
    HiddenFromAddressListsEnabled = $false
}
# Группа доступа Jira, в которую админ уже вложил дубль.
$global:ADStore.Groups += [PSCustomObject]@{
    DistinguishedName = 'CN=JIRA-SPACE-SALES,OU=Jira,DC=transitcard,DC=ru'
    Name = 'JIRA-SPACE-SALES'; SamAccountName = 'JIRA-SALES'; DisplayName = 'JIRA-SPACE-SALES'
    Description = 'Доступ к пространству Jira'
    Member = @("CN=Отдел продаж и маркетинга,$OU"); MemberOf = @(); ManagedBy = $null
    mail = $null; info = $null; whenCreated = (Get-Date).AddDays(-100)
    objectSid = [PSCustomObject]@{ Value = 'S-1-5-21-1-1-1-777' }
    HiddenFromAddressListsEnabled = $false
}
(Get-G 'Отдел продаж и маркетинга').MemberOf = @('CN=JIRA-SPACE-SALES,OU=Jira,DC=transitcard,DC=ru')
Set-Excel @(
    @{ Child = 'Отдел продаж и маркетинга'; Parent = 'Коммерческий блок'; Manager = '1001' },
    @{ Child = 'Отдел закупок';             Parent = 'Коммерческий блок'; Manager = '2001' }
)
Move-Users -Sams @('sales1','sales2','sales3','sales4','sales5','sales6') -ToDept 'Отдел продаж и маркетинга'
Invoke-Sync
$surv = Get-G 'Отдел продаж и маркетинга'
Check 'выжила старая группа со своим SID' ($surv -ne $null -and $surv.objectSid.Value -eq $sidOld)
Check 'дубль отставлен под именем ZZ-Дубль-*' (@($global:ADStore.Groups | Where-Object { $_.Name -like 'ZZ-Дубль-*' }).Count -eq 1)
$jira = $global:ADStore.Groups | Where-Object { $_.Name -eq 'JIRA-SPACE-SALES' }
Check 'доступ Jira перенесён на выжившую группу' (@($jira.Member) -contains $surv.DistinguishedName)
Check 'состав выжившей группы полный (12)' ((Get-MemSams 'Отдел продаж и маркетинга').Count -eq 12) "= $((Get-MemSams 'Отдел продаж и маркетинга').Count)"

''
'============================================================'
'ТЕСТ 9. Предохранители: битая выгрузка и пропавший Department'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
# Выгрузка 1С «схлопнулась»: остался один отдел из трёх.
Set-Excel @( @{ Child = 'Отдел закупок'; Parent = 'Коммерческий блок'; Manager = '2001' } )
Invoke-Sync
Check 'состав продаж не тронут при пропаже отдела из 1С' ((Get-MemSams 'Отдел продаж').Count -eq 12) "= $((Get-MemSams 'Отдел продаж').Count)"

Reset-Store; Build-BaseOrg; Invoke-Sync
# У всех сотрудников продаж пропал Department (сбой выгрузки в AD).
foreach ($u in $global:ADStore.Users) { if ($u.Department -eq 'Отдел продаж') { $u.Department = $null } }
Invoke-Sync
Check 'живая группа не опустошена при пропаже Department' ((Get-MemSams 'Отдел продаж').Count -eq 12) "= $((Get-MemSams 'Отдел продаж').Count)"
Check 'сработал предохранитель пустого состава' ($script:LastOutput -match 'целевой состав пуст')

''
'============================================================'
'ТЕСТ 10. Общий предохранитель массового исключения'
'============================================================'
Reset-Store
Set-Excel @( @{ Child = 'Отдел продаж'; Parent = 'Коммерческий блок'; Manager = '1001' } )
Add-User -Sam 'boss.sales' -Name 'Руководитель Продаж' -Dept 'Отдел продаж' -EmpNum '1001'
for ($i = 1; $i -le 20; $i++) { Add-User -Sam "sales$i" -Name "Продавец $i" -Dept 'Отдел продаж' -EmpNum "10$i" }
Invoke-Sync
Check 'исходно 21 человек' ((Get-MemSams 'Отдел продаж').Count -eq 21) "= $((Get-MemSams 'Отдел продаж').Count)"
# Половину сотрудников «потеряли»: у них сменился Department на существующую вершину.
Move-Users -Sams (1..12 | ForEach-Object { "sales$_" }) -ToDept 'Коммерческий блок'
Invoke-Sync
Check 'массовое исключение заблокировано' ($script:LastOutput -match 'АВАРИЯ')
Check 'состав не изменён' ((Get-MemSams 'Отдел продаж').Count -eq 21) "= $((Get-MemSams 'Отдел продаж').Count)"

''
'============================================================'
'ТЕСТ 11. Возврат к прежнему названию — привязка по реестру'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$sid0 = (Get-G 'Отдел продаж').objectSid.Value
Set-Excel @(
    @{ Child = 'Отдел продаж и маркетинга'; Parent = 'Коммерческий блок'; Manager = '1001' },
    @{ Child = 'Отдел закупок';             Parent = 'Коммерческий блок'; Manager = '2001' }
)
Move-Users -Sams (@('boss.sales') + (1..11 | ForEach-Object { "sales$_" })) -ToDept 'Отдел продаж и маркетинга'
Invoke-Sync
Check 'первое переименование прошло' ((Get-G 'Отдел продаж и маркетинга').objectSid.Value -eq $sid0)
# Передумали: вернули прежнее название, люди ещё не переехали обратно.
Set-Excel @(
    @{ Child = 'Отдел продаж';  Parent = 'Коммерческий блок'; Manager = '1001' },
    @{ Child = 'Отдел закупок'; Parent = 'Коммерческий блок'; Manager = '2001' }
)
Invoke-Sync
Check 'возврат к прежнему названию распознан по реестру' ((Get-G 'Отдел продаж') -ne $null -and (Get-G 'Отдел продаж').objectSid.Value -eq $sid0)
Check 'новая группа не создана' (@($global:ADStore.Groups | Where-Object { $_.Description -eq $Desc }).Count -eq 3)

''
'============================================================'
'ТЕСТ 12. Защищённая группа не уходит в удаление'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
$mail = (Get-G 'Отдел продаж').mail
Set-Content -Path (Join-Path $Work 'protected_groups.json') -Value (ConvertTo-Json @(@{ email = $mail })) -Encoding UTF8
Set-Excel @( @{ Child = 'Отдел закупок'; Parent = 'Коммерческий блок'; Manager = '2001' } )
Move-Users -Sams (@('boss.sales') + (1..11 | ForEach-Object { "sales$_" })) -ToDept 'Отдел закупок'
Invoke-Sync
(Get-G 'Отдел продаж').info = "QUARANTINE:{0:yyyy-MM-dd}" -f (Get-Date).AddDays(-90)
Invoke-Sync
Check 'защищённая группа жива' ($null -ne (Get-G 'Отдел продаж'))
Check 'защищённая группа не переименована' (@($global:ADStore.Groups | Where-Object { $_.Name -eq 'Отдел продаж' }).Count -eq 1)
Check 'защищённая группа не в карантине' ((Get-G 'Отдел продаж') -and -not (Get-G 'Отдел продаж').HiddenFromAddressListsEnabled)
Check 'состав защищённой группы не тронут' ((Get-MemSams 'Отдел продаж').Count -eq 12) "= $((Get-MemSams 'Отдел продаж').Count)"

''
'============================================================'
'ТЕСТ 13. Отключённый сотрудник исключается из группы'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
Check 'до увольнения 12' ((Get-MemSams 'Отдел продаж').Count -eq 12)
($global:ADStore.Users | Where-Object { $_.SamAccountName -eq 'sales1' }).Enabled = $false
Invoke-Sync
Check 'уволенный исключён' (-not ((Get-MemSams 'Отдел продаж') -contains 'sales1'))
Check 'осталось 11' ((Get-MemSams 'Отдел продаж').Count -eq 11) "= $((Get-MemSams 'Отдел продаж').Count)"

''
'============================================================'
'ТЕСТ 14. Сотрудник без табельного номера в Users.Corp'
'============================================================'
Reset-Store; Build-BaseOrg
Add-User -Sam 'nonum' -Name 'Без Табельного' -Dept 'Отдел продаж' -EmpNum $null
Invoke-Sync
Check 'сотрудник без EmployeeNumber попал в группу' ((Get-MemSams 'Отдел продаж') -contains 'nonum')
Check 'всего 13' ((Get-MemSams 'Отдел продаж').Count -eq 13) "= $((Get-MemSams 'Отдел продаж').Count)"

''
'============================================================'
'ТЕСТ 15. Истёкший псевдоним снимается'
'============================================================'
Reset-Store; Build-BaseOrg; Invoke-Sync
Set-Content -Path (Join-Path $Work 'department_aliases.json') -Encoding UTF8 -Value (ConvertTo-Json @(
    @{ Old = 'Старое имя'; New = 'Отдел продаж'; Since = '2020-01-01'; Expires = '2020-03-01' },
    @{ Old = 'Живое имя';  New = 'Отдел продаж'; Since = '2020-01-01'; Expires = (Get-Date).AddDays(30).ToString('yyyy-MM-dd') }
))
Invoke-Sync
$al = Read-Json 'department_aliases.json'
Check 'истёкший псевдоним удалён'  (@($al | Where-Object { $_.Old -eq 'Старое имя' }).Count -eq 0)
Check 'действующий псевдоним остался' (@($al | Where-Object { $_.Old -eq 'Живое имя' }).Count -eq 1)

''
'============================================================'
"ИТОГО: успешно $script:Pass, провалено $script:Fail"
if ($script:Fail -gt 0) { 'Провалены:'; $script:Failures | ForEach-Object { "  - $_" } }
'============================================================'
