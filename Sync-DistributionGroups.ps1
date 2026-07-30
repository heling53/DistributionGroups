#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Синхронизация групп рассылки AD/Exchange со справочником подразделений 1С (Подразделения.xlsx).
.DESCRIPTION
    Единственный источник данных о подразделениях — Подразделения.xlsx (выгрузка 1С).
    Из него берутся:
      - уникальные подразделения  — для создания групп рассылки (колонки objFullName/objName);
      - иерархия (parent/child)    — для вложенности групп (objName -> objFullName);
      - руководители подразделений — колонка «Руководитель подразделения» (табельный номер).
    Дополнительно:
      - protected_groups.json — защита от удаления.
    Пользователи берутся из AD и распределяются по группам по названию подразделения.
.NOTES
    Все операции выполняются на одном DC ($DcServer), read-after-write согласован.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$TargetOU                = 'OU=Группы рассылки,DC=transitcard,DC=ru',
    [string]$GroupDesc               = 'Группа рассылки подразделения_v1.0',
    [string]$MailDomain              = 'transitcard.ru',
    [string]$ExchangeServer          = 'srv-ex16-01.transitcard.ru',
    [string]$DcServer                = 'srv-dc3.transitcard.ru',
    [string]$ProtectedGroupsJsonPath = 'D:\Scripts\Группы рассылки\protected_groups.json',
    [string]$ExcelPath               = 'D:\Scripts\verify1C\1c\Подразделения.xlsx',
    [string]$LogPath                 = 'D:\Scripts\Группы рассылки\Logs',
    [int]$SamMaxLength               = 20,
    [string]$UsersCorpOU             = 'OU=Users.Corp,DC=transitcard,DC=ru',
    # Кастомная верхнеуровневая структура. «Руководство» приходит из 1С, «Все сотрудники» — нет
    # (создаётся один раз скриптом New-AllEmployeesGroup.ps1 с адресом employees@<MailDomain>).
    [string]$LeadershipGroupName     = 'Руководство',
    [string]$AllEmployeesGroupName   = 'Все сотрудники'
)

$ErrorActionPreference = 'Stop'

# ==========================================
# Кодировка консоли (только в интерактивном режиме)
# ==========================================
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
} catch {
    Write-Warning "Не удалось установить кодировку консоли (возможно, скрипт запущен без консольного окна)."
}
$PSDefaultParameterValues['Get-Content:Encoding'] = 'UTF8'

# ==========================================
# Логирование
# ==========================================
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
$transcript = Join-Path $LogPath ("DG-Sync_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $transcript -Append | Out-Null

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) {
        'ERROR' { 'Red' }; 'WARN' { 'Yellow' }; 'OK' { 'Green' }
        'STEP'  { 'Cyan' }; 'DEBUG' { 'DarkGray' }; default { 'Gray' }
    }
    Write-Host ("[{0:HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message) -ForegroundColor $color
}

# ==========================================
# Транслитерация (словарь — один раз на сессию)
# ==========================================
$script:CyrLatMap = @{}
$low = @{'а'='a';'б'='b';'в'='v';'г'='g';'д'='d';'е'='e';'ё'='e';'ж'='zh';'з'='z';'и'='i';'й'='y';'к'='k';'л'='l';'м'='m';'н'='n';'о'='o';'п'='p';'р'='r';'с'='s';'т'='t';'у'='u';'ф'='f';'х'='h';'ц'='ts';'ч'='ch';'ш'='sh';'щ'='sch';'ъ'='';'ы'='y';'ь'='';'э'='e';'ю'='yu';'я'='ya'}
$up  = @{'А'='A';'Б'='B';'В'='V';'Г'='G';'Д'='D';'Е'='E';'Ё'='E';'Ж'='Zh';'З'='Z';'И'='I';'Й'='Y';'К'='K';'Л'='L';'М'='M';'Н'='N';'О'='O';'П'='P';'Р'='R';'С'='S';'Т'='T';'У'='U';'Ф'='F';'Х'='H';'Ц'='Ts';'Ч'='Ch';'Ш'='Sh';'Щ'='Sch';'Ъ'='';'Ы'='Y';'Ь'='';'Э'='E';'Ю'='Yu';'Я'='Ya'}
foreach ($k in $low.Keys) { $script:CyrLatMap[[char]$k] = $low[$k] }
foreach ($k in $up.Keys)  { $script:CyrLatMap[[char]$k] = $up[$k] }

function Convert-CyrillicToLatin {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $sb = [System.Text.StringBuilder]::new($Text.Length)
    foreach ($ch in $Text.ToCharArray()) {
        if ($script:CyrLatMap.ContainsKey($ch)) { [void]$sb.Append($script:CyrLatMap[$ch]) }
        else { [void]$sb.Append($ch) }
    }
    return ($sb.ToString() -replace '[^a-zA-Z0-9\-_]', '')
}

function Get-Initials {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($word in ($Text -split '[\s\-]+')) {
        $clean = $word -replace '[^\p{L}0-9]', ''
        if ($clean.Length -gt 0) { [void]$sb.Append($clean[0]) }
    }
    return $sb.ToString()
}

function Normalize-Name {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return ($Text -replace '\s+', ' ').Trim().ToLowerInvariant()
}

function Get-UniqueSamAccountName {
    param([string]$BaseName, [hashtable]$Used, [int]$MaxLen = 20)
    $base = $BaseName
    if ($base.Length -gt $MaxLen) { $base = $base.Substring(0, $MaxLen) }
    $candidate = $base; $i = 1
    while ($Used.ContainsKey($candidate) -or
           (Get-ADGroup -Filter "SamAccountName -eq '$candidate'" -Server $DcServer -ErrorAction SilentlyContinue)) {
        $suffix = [string]$i
        $cut = $MaxLen - $suffix.Length
        if ($cut -lt 1) { throw "Не удалось подобрать SamAccountName для '$BaseName'" }
        $candidate = $base.Substring(0, [Math]::Min($base.Length, $cut)) + $suffix
        $i++
    }
    return $candidate
}

# ==========================================
# Подключение к Exchange
# ==========================================
$session = $null
try {
    Write-Log "Подключение к Exchange $ExchangeServer..." 'STEP'
    $session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeServer) `
                             -Authentication Kerberos
    Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
    Write-Log "Сессия Exchange открыта." 'OK'

    # ----------------------------------------
    # Загрузка защищённых групп
    # ----------------------------------------
    $protectedByEmail = @{}
    if (Test-Path $ProtectedGroupsJsonPath) {
        try {
            $protectedList = Get-Content -Path $ProtectedGroupsJsonPath -Raw | ConvertFrom-Json
            foreach ($entry in $protectedList) {
                if ($entry.email) {
                    $protectedByEmail[$entry.email.ToLowerInvariant().Trim()] = $entry
                }
            }
            Write-Log "Загружено $($protectedByEmail.Count) защищённых групп." 'OK'
        } catch {
            Write-Log "Ошибка чтения protected_groups.json: $($_.Exception.Message)" 'ERROR'
        }
    } else {
        Write-Log "protected_groups.json не найден — защита от удаления отключена." 'WARN'
    }

    # ==========================================
    # Загрузка справочника 1С (Подразделения.xlsx) — единственный источник.
    # Отсюда: уникальные отделы, иерархия, руководители, ID для аудита.
    # ==========================================
    Write-Log "Загрузка справочника 1С: $ExcelPath" 'STEP'
    if (-not (Test-Path $ExcelPath)) {
        throw "Справочник подразделений не найден: $ExcelPath"
    }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "Модуль ImportExcel не установлен. Установите: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel

    $excelHierarchy = New-Object System.Collections.Generic.List[object]
    $excelManagers  = New-Object System.Collections.Generic.List[object]
    $departmentList = New-Object System.Collections.Generic.List[string]
    $seenDept       = @{}   # нормализованное имя -> $true (дедупликация)

    $excelData = Import-Excel -Path $ExcelPath -StartRow 2
    foreach ($row in $excelData) {
        $childName  = if ($row.objFullName) { ($row.objFullName.ToString() -replace '\s+', ' ').Trim() } else { $null }
        $parentName = if ($row.objName)     { ($row.objName.ToString()     -replace '\s+', ' ').Trim() } else { $null }
        # Срезаем ведущие нули у EmployeeNumber руководителя (Excel-формат "0000003966").
        # Исключение: табельный номер "0000-00013" — единственный с дефисом, ведущие нули
        # значимы (в AD хранится как "0000-00013"). Для него значение оставляем как есть.
        $managerEN  = if ($row.'Руководитель подразделения') {
            $en = $row.'Руководитель подразделения'.ToString().Trim()
            if ($en -eq '0000-00013') { $en } else { $en -replace '^0+', '' }
        } else { $null }

        if ($childName -and $parentName -and ($childName -ne $parentName)) {
            $excelHierarchy.Add([PSCustomObject]@{ Child = $childName; Parent = $parentName })
        }
        if ($childName -and $managerEN) {
            $excelManagers.Add([PSCustomObject]@{ Department = $childName; EmployeeNumber = $managerEN })
        }

        # Уникальные подразделения: и сам узел (objFullName), и его родитель (objName),
        # чтобы для любого уровня иерархии существовала группа.
        foreach ($name in @($childName, $parentName)) {
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $norm = Normalize-Name $name
                if (-not $seenDept.ContainsKey($norm)) {
                    $seenDept[$norm] = $true
                    $departmentList.Add($name)
                }
            }
        }
    }

    $departments = $departmentList | Sort-Object
    Write-Log ("1С: отделов={0}, иерархия={1}, руководителей={2}" -f `
        @($departments).Count, $excelHierarchy.Count, $excelManagers.Count) 'OK'

    if (@($departments).Count -eq 0) {
        throw "Из $ExcelPath не получено ни одного подразделения (проверьте колонки objFullName/objName и -StartRow)."
    }

    # Набор имён групп, которые ДОЛЖНЫ существовать по справочнику 1С.
    # Ключи нормализованы — так же, как имена создаваемых групп,
    # чтобы Шаг 4 не удалял пустые группы, реально присутствующие в 1С.
    $excelGroupKeys = @{}
    foreach ($d in $departments) {
        $excelGroupKeys[(Normalize-Name $d)] = $true
    }
    # «Все сотрудники» — кастомная верхнеуровневая группа, которой нет в выгрузке 1С.
    # Добавляем её в набор «обязательно существующих», чтобы Шаг 4 никогда её не удалял
    # (даже если в какой-то момент она временно окажется пустой).
    $excelGroupKeys[(Normalize-Name $AllEmployeesGroupName)] = $true

    # ==========================================
    # ШАГ 1. Создание групп из справочника 1С
    # ==========================================
    Write-Log "Шаг 1: создание групп из справочника 1С" 'STEP'

    # Один запрос: все управляемые группы в OU
    $existingGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                  -SearchBase $TargetOU -Server $DcServer `
                                  -Properties Name, SamAccountName, mail, DisplayName, Member
    $groupsByName = @{}
    $usedSam      = @{}
    foreach ($g in $existingGroups) {
        $groupsByName[$g.Name.Trim()] = $g
        $usedSam[$g.SamAccountName]  = $true
    }

    foreach ($dept in $departments) {
        $groupName = $dept
        if ($groupsByName.ContainsKey($groupName)) { continue }

        $latin = Convert-CyrillicToLatin -Text (Get-Initials -Text $dept)
        $base  = ('DG_' + $latin).ToUpper()
        if ([string]::IsNullOrWhiteSpace($base)) { $base = 'DG' }

        try {
            $sam = Get-UniqueSamAccountName -BaseName $base -Used $usedSam -MaxLen $SamMaxLength
        } catch {
            Write-Log "Ошибка подбора SamAccountName для '$dept': $_" 'ERROR'
            continue
        }
        $usedSam[$sam] = $true

        if ($PSCmdlet.ShouldProcess($groupName, "New-ADGroup + Enable-DistributionGroup")) {
            try {
                New-ADGroup -Name $groupName -SamAccountName $sam `
                            -GroupCategory Distribution -GroupScope Universal `
                            -Path $TargetOU -Description $GroupDesc -DisplayName $groupName `
                            -Server $DcServer

                Enable-DistributionGroup -Identity $sam -Alias $sam `
                    -PrimarySmtpAddress ("{0}@{1}" -f $sam.ToLower(), $MailDomain) `
                    -DomainController $DcServer

                Set-DistributionGroup -Identity $sam `
                    -RequireSenderAuthenticationEnabled $false `
                    -DomainController $DcServer

                Write-Log "Создана группа: $groupName ($sam)" 'OK'
            } catch {
                Write-Log "Ошибка создания '$groupName': $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 1.1. ПОЛНАЯ ОЧИСТКА СОСТАВА ВСЕХ ГРУПП
    # ==========================================
    Write-Log "Шаг 1.1: полная очистка состава всех групп" 'STEP'
    $allGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" -SearchBase $TargetOU -Server $DcServer
    foreach ($g in $allGroups) {
        if ($PSCmdlet.ShouldProcess($g.Name, "Clear members")) {
            try {
                Set-ADGroup -Identity $g -Clear Member -Server $DcServer
            } catch {
                Write-Log "Ошибка очистки $($g.Name): $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # Перечитываем актуальный список (группы пусты)
    $managedGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                 -SearchBase $TargetOU -Server $DcServer `
                                 -Properties Member, mail
    $managedGroupsMap = @{}
    foreach ($g in $managedGroups) {
        $managedGroupsMap[(Normalize-Name $g.Name)] = $g
    }

    # ==========================================
    # ШАГ 2. Наполнение пользователями (только добавление)
    # ==========================================
    Write-Log "Шаг 2: добавление пользователей" 'STEP'

    $filterUsers = @"
Department -like '*'
  -and Enabled -eq 'True'
  -and EmployeeNumber -like '*'
  -and SamAccountName -notlike 'adm-*'
  -and SamAccountName -notlike 'vpn-*'
  -and SamAccountName -notlike '*test*'
  -and Name -notlike '*ADM*'
  -and Name -notlike '*VPN*'
  -and Name -notlike '*KZ*'
"@ -replace "`r`n", " " -replace "`n", " "

    $users = Get-ADUser -Filter $filterUsers `
                        -Properties Department, EmployeeNumber, Name `
                        -ResultPageSize 1000 -Server $DcServer

    $usersByDept   = @{}
    $usersByEmpNum = @{}
    foreach ($u in $users) {
        $key = Normalize-Name $u.Department
        if (-not $usersByDept.ContainsKey($key)) {
            $usersByDept[$key] = New-Object System.Collections.Generic.List[object]
        }
        $usersByDept[$key].Add($u)
        if ($u.EmployeeNumber) {
            $usersByEmpNum[$u.EmployeeNumber.ToString().Trim()] = $u
        }
    }

    foreach ($g in $managedGroups) {
        # Точное совпадение по имени подразделения: имена групп берутся из 1С
        # и совпадают с атрибутом Department в AD.
        $key = Normalize-Name $g.Name
        if (-not $usersByDept.ContainsKey($key)) { continue }
        # .ToArray() приводит List[object] к object[]: дальше работаем с обычным массивом.
        $desired = $usersByDept[$key].ToArray()
        if ($desired.Count -eq 0) { continue }
        $desiredDNs = @($desired | ForEach-Object { $_.DistinguishedName })
        $toAdd = $desiredDNs | Where-Object { $_ -and ($g.Member -notcontains $_) }
        if ($toAdd.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($g.Name, "Add $($toAdd.Count) users")) {
                try {
                    Add-ADGroupMember -Identity $g -Members $toAdd -Server $DcServer
                    Write-Log "+$($toAdd.Count) -> $($g.Name)" 'OK'
                } catch { Write-Log "Add fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
            }
        }
    }

    # ==========================================
    # ШАГ 2.1. Добавление руководителей подразделений из 1С
    # ==========================================
    Write-Log "Шаг 2.1: руководители подразделений" 'STEP'
    foreach ($m in $excelManagers) {
        if (-not $usersByEmpNum.ContainsKey($m.EmployeeNumber)) {
            Write-Log "Руководитель EmployeeNumber=$($m.EmployeeNumber) для '$($m.Department)' не найден в AD" 'WARN'
            continue
        }
        $manager = $usersByEmpNum[$m.EmployeeNumber]
        $deptKey = Normalize-Name $m.Department
        if (-not $managedGroupsMap.ContainsKey($deptKey)) {
            Write-Log "Группа подразделения '$($m.Department)' не найдена" 'WARN'
            continue
        }
        $group = $managedGroupsMap[$deptKey]
        if ($PSCmdlet.ShouldProcess($group.Name, "Add manager $($manager.SamAccountName)")) {
            try {
                Add-ADGroupMember -Identity $group -Members $manager.DistinguishedName -Server $DcServer -ErrorAction Stop
                Write-Log "Руководитель $($manager.SamAccountName) -> $($group.Name)" 'OK'
            } catch {
                if ($_.Exception.Message -notmatch 'already a member') {
                    Write-Log "Add manager fail $($group.Name): $($_.Exception.Message)" 'ERROR'
                }
            }
        }
        if ($PSCmdlet.ShouldProcess($group.Name, "Set ManagedBy=$($manager.SamAccountName)")) {
            try {
                Set-ADGroup -Identity $group -ManagedBy $manager.DistinguishedName -Server $DcServer
                Write-Log "ManagedBy: $($group.Name) <- $($manager.SamAccountName)" 'OK'
            } catch {
                Write-Log "Set ManagedBy fail $($group.Name): $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 3. Иерархия из 1С (Excel)
    # ==========================================
    # Подмена маппинга «Руководство» <-> «Все сотрудники»:
    #   «Руководство» больше НЕ верхнеуровневая группа, в которую входят все. Везде, где
    #   1С указывает родителем (вершиной) «Руководство», подставляем кастомную группу
    #   «Все сотрудники» (employees@), которой в выгрузке 1С нет. В итоге бывшие прямые
    #   подчинённые «Руководство» вкладываются в «Все сотрудники», а саму «Руководство»
    #   отдельным ребром вкладываем в «Все сотрудники» в конце шага.
    Write-Log "Шаг 3: иерархия" 'STEP'

    $leadershipKey   = Normalize-Name $LeadershipGroupName
    $allEmployeesKey = Normalize-Name $AllEmployeesGroupName
    if (-not $managedGroupsMap.ContainsKey($allEmployeesKey)) {
        Write-Log ("Кастомная группа '{0}' не найдена в {1}. Создайте её один раз скриптом New-AllEmployeesGroup.ps1 (employees@{2}). Верхний уровень иерархии построен не будет." -f `
            $AllEmployeesGroupName, $TargetOU, $MailDomain) 'ERROR'
    }

    foreach ($row in $excelHierarchy) {
        $childKey  = Normalize-Name $row.Child
        $parentKey = Normalize-Name $row.Parent

        # «Руководство» вкладываем только в «Все сотрудники» (ниже, явным ребром),
        # поэтому её собственные родительские связи из 1С здесь игнорируем.
        if ($childKey -eq $leadershipKey) { continue }

        # Подмена родителя: «Руководство» -> «Все сотрудники».
        $parentName = $row.Parent
        if ($parentKey -eq $leadershipKey) {
            $parentKey  = $allEmployeesKey
            $parentName = $AllEmployeesGroupName
        }

        if (-not $managedGroupsMap.ContainsKey($parentKey)) { Write-Log "Родитель '$parentName' не найден" 'WARN'; continue }
        if (-not $managedGroupsMap.ContainsKey($childKey))  { Write-Log "Потомок '$($row.Child)' не найден"  'WARN'; continue }
        $parent = $managedGroupsMap[$parentKey]
        $child  = $managedGroupsMap[$childKey]
        try {
            if ($PSCmdlet.ShouldProcess("$($child.Name) -> $($parent.Name)", "Add child group")) {
                Add-ADGroupMember -Identity $parent -Members $child -Server $DcServer -ErrorAction Stop
                Write-Log "Иерархия: $($child.Name) -> $($parent.Name)" 'OK'
            }
        } catch {
            if ($_.Exception.Message -notmatch 'already a member') {
                Write-Log "Иерархия fail: $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # Саму «Руководство» вкладываем в «Все сотрудники» (явное ребро верхнего уровня).
    if ($managedGroupsMap.ContainsKey($leadershipKey) -and $managedGroupsMap.ContainsKey($allEmployeesKey)) {
        $leadershipGroup   = $managedGroupsMap[$leadershipKey]
        $allEmployeesGroup = $managedGroupsMap[$allEmployeesKey]
        try {
            if ($PSCmdlet.ShouldProcess("$($leadershipGroup.Name) -> $($allEmployeesGroup.Name)", "Add child group")) {
                Add-ADGroupMember -Identity $allEmployeesGroup -Members $leadershipGroup -Server $DcServer -ErrorAction Stop
                Write-Log "Иерархия: $($leadershipGroup.Name) -> $($allEmployeesGroup.Name)" 'OK'
            }
        } catch {
            if ($_.Exception.Message -notmatch 'already a member') {
                Write-Log "Иерархия fail ($($leadershipGroup.Name) -> $($allEmployeesGroup.Name)): $($_.Exception.Message)" 'ERROR'
            }
        }
    } elseif (-not $managedGroupsMap.ContainsKey($leadershipKey)) {
        Write-Log "Группа '$LeadershipGroupName' не найдена — не удалось вложить её в '$AllEmployeesGroupName'" 'WARN'
    }

    # ==========================================
    # ШАГ 3.2. Исключение: пользователи OU Users.Corp без EmployeeNumber
    # ==========================================
    # Шаг 2 жёстко требует EmployeeNumber, поэтому сотрудники без табельного номера
    # в группы не попадают. Здесь, перед удалением пустых групп, добавляем исключение:
    # пользователей из $UsersCorpOU, у которых всё остальное в порядке (включён,
    # есть Department, не попадает под adm/vpn/test/KZ-исключения), но нет EmployeeNumber.
    Write-Log "Шаг 3.2: исключение — пользователи $UsersCorpOU без EmployeeNumber" 'STEP'

    $filterNoEmpNum = @"
Department -like '*'
  -and Enabled -eq 'True'
  -and -not (EmployeeNumber -like '*')
  -and SamAccountName -notlike 'adm-*'
  -and SamAccountName -notlike 'vpn-*'
  -and SamAccountName -notlike '*test*'
  -and Name -notlike '*ADM*'
  -and Name -notlike '*VPN*'
  -and Name -notlike '*KZ*'
"@ -replace "`r`n", " " -replace "`n", " "

    $exceptionUsers = @()
    try {
        $exceptionUsers = Get-ADUser -Filter $filterNoEmpNum `
                                     -SearchBase $UsersCorpOU `
                                     -Properties Department, Name `
                                     -ResultPageSize 1000 -Server $DcServer
    } catch {
        Write-Log "Не удалось получить пользователей из ${UsersCorpOU}: $($_.Exception.Message)" 'ERROR'
    }

    $exUsersByDept = @{}
    foreach ($u in $exceptionUsers) {
        $key = Normalize-Name $u.Department
        if (-not $exUsersByDept.ContainsKey($key)) {
            $exUsersByDept[$key] = New-Object System.Collections.Generic.List[object]
        }
        $exUsersByDept[$key].Add($u)
    }
    Write-Log ("Исключений найдено: {0} пользователей, {1} подразделений" -f `
        @($exceptionUsers).Count, $exUsersByDept.Count) 'OK'

    if ($exUsersByDept.Count -gt 0) {
        # Перечитываем актуальный состав групп (после шагов 2/2.1/3), чтобы не дублировать.
        $groupsForExcept = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                       -SearchBase $TargetOU -Server $DcServer `
                                       -Properties Member
        foreach ($g in $groupsForExcept) {
            $key = Normalize-Name $g.Name
            if (-not $exUsersByDept.ContainsKey($key)) { continue }
            $desired = $exUsersByDept[$key].ToArray()
            if ($desired.Count -eq 0) { continue }
            $usersToAdd = @($desired | Where-Object { $_.DistinguishedName -and ($g.Member -notcontains $_.DistinguishedName) })
            if ($usersToAdd.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess($g.Name, "Add $($usersToAdd.Count) users без EmployeeNumber")) {
                    try {
                        Add-ADGroupMember -Identity $g -Members @($usersToAdd.DistinguishedName) -Server $DcServer
                        Write-Log "+$($usersToAdd.Count) (Users.Corp без EmpNum) -> $($g.Name)" 'OK'
                        foreach ($u in $usersToAdd) {
                            Write-Log "    $($u.SamAccountName) — $($u.Name)" 'OK'
                        }
                    } catch { Write-Log "Add fail (без EmpNum) $($g.Name): $($_.Exception.Message)" 'ERROR' }
                }
            }
        }
    }

    # ==========================================
    # ШАГ 4. Удаление групп: только если ПУСТАЯ И отсутствует в справочнике 1С
    # ==========================================
    Write-Log "Шаг 4: удаление пустых групп, отсутствующих в 1С" 'STEP'
    $finalGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                               -SearchBase $TargetOU -Server $DcServer `
                               -Properties Member, mail
    foreach ($g in $finalGroups) {
        # 1) Не трогаем непустые группы
        if ($g.Member -and $g.Member.Count -gt 0) { continue }

        # 2) Не удаляем, если подразделение есть в справочнике 1С (даже если группа сейчас пуста,
        #    например новый отдел без сотрудников). Точная сверка по нормализованному имени.
        if ($excelGroupKeys.ContainsKey((Normalize-Name $g.Name))) {
            Write-Log "Есть в 1С, пропуск удаления пустой группы: $($g.Name)" 'WARN'
            continue
        }

        # 3) Не удаляем защищённые в protected_groups.json
        $email = if ($g.mail) { $g.mail.ToLowerInvariant().Trim() } else { $null }
        if ($email -and $protectedByEmail.ContainsKey($email)) {
            Write-Log "Защищена JSON, пропуск: $($g.Name) ($email)" 'WARN'
            continue
        }
        if ($PSCmdlet.ShouldProcess($g.Name, "Disable + Remove")) {
            try {
                Disable-DistributionGroup -Identity $g.DistinguishedName -Confirm:$false -DomainController $DcServer
            } catch { Write-Log "Disable fail $($g.Name): $($_.Exception.Message)" 'WARN' }
            try {
                Remove-ADGroup -Identity $g -Confirm:$false -Server $DcServer
                Write-Log "Удалена пустая группа: $($g.Name)" 'OK'
            } catch { Write-Log "Remove fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
        }
    }
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
    throw
}
finally {
    if ($session) {
        Remove-PSSession $session -ErrorAction SilentlyContinue
        Write-Log "Сессия Exchange закрыта." 'OK'
    }
    Write-Log "Готово (DC=$DcServer)." 'STEP'
    Stop-Transcript | Out-Null
}
