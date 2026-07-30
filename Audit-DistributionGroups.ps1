# ========== АБСОЛЮТНЫЕ ПУТИ К ФАЙЛАМ ==========
$TemplatePath = "D:\Scripts\DistributionGroups\template_message1.html"
$LogFile      = "D:\Scripts\DistributionGroups\task.log"
$ExcelPath    = "D:\Scripts\verify1C\1c\Подразделения.xlsx"
$ExcelSheet   = "Лист_1"

# Кастомная верхнеуровневая структура (синхронно со Sync-DistributionGroups.ps1):
# «Руководство» приходит из 1С, но больше не является корнем — она вложена в кастомную
# группу «Все сотрудники», которой в выгрузке 1С нет.
$LeadershipGroupName   = "Руководство"
$AllEmployeesGroupName = "Все сотрудники"

# CSV-файлы
$Csv_NonCorrADGroup        = "D:\Scripts\DistributionGroups\NonCorrADGroup.csv"
$Csv_NCMUpload             = "D:\Scripts\DistributionGroups\NCMUpload.csv"
$Csv_NonCorrUserMembership = "D:\Scripts\DistributionGroups\NonCorrUserMembership.csv"
$Csv_UserDepToAdd          = "D:\Scripts\DistributionGroups\UserDepToAdd.csv"
$Csv_UsersToExclude        = "D:\Scripts\DistributionGroups\UsersToExclude.csv"
$Csv_DGroupsToDel          = "D:\Scripts\DistributionGroups\DGroupsToDel.csv"

# Защищённые группы (тот же файл, что и в скрипте синхронизации)
$ProtectedGroupsJsonPath   = "D:\Scripts\Группы рассылки\protected_groups.json"

# Параметры почты
$MailServer    = "owa.transitcard.ru"
$MailPort      = 25
$MailSender    = "infosec-notifications@pprcard.ru"
$MailRecipient = @("itsecurity@pprcard.ru")
$MailSubject   = "[Группы рассылки] Менеджмент групп рассылки подразделений"

# ========== ПРОВЕРКА ПОВТОРНОГО ЗАПУСКА ==========
$mutex = New-Object System.Threading.Mutex($false, "Global\DistributionGroupsScript")
if (-not $mutex.WaitOne(0)) {
    $date = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LogFile -Value "$date [Warning] Скрипт уже выполняется. Выход."
    exit
}

# ========== ФУНКЦИЯ ЛОГИРОВАНИЯ ==========
function Out-Log {
    param([string]$Message)
    $date = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LogFile -Value "$date $Message"
}

# ========== ПРОВЕРКА НАЛИЧИЯ МОДУЛЕЙ ==========
$missingModules = @()
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    $missingModules += "ActiveDirectory"
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    $missingModules += "ImportExcel"
}
if ($missingModules.Count -gt 0) {
    $errMsg = "Отсутствуют требуемые модули: $($missingModules -join ', ')"
    Out-Log "[Error] $errMsg"
    Write-Error $errMsg
    exit 1
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Out-Log "[Information] Модуль ActiveDirectory загружен."
} catch {
    Out-Log "[Error] Не удалось загрузить модуль ActiveDirectory: $($_.Exception.Message)"
    exit 1
}

try {
    Import-Module ImportExcel -ErrorAction Stop
    Out-Log "[Information] Модуль ImportExcel загружен."
} catch {
    Out-Log "[Error] Не удалось загрузить модуль ImportExcel: $($_.Exception.Message)"
    exit 1
}

# ========== ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ (ТОЧНЫЙ ПОИСК ГРУППЫ) ==========
Function Find-GroupMatch {
    Param([string]$SearchName, [hashtable]$Map)

    if ([string]::IsNullOrWhiteSpace($SearchName)) { return $null }
    if ($Map.ContainsKey($SearchName)) { return $Map[$SearchName] }
    return $null
}

# ========== ЗАГРУЗКА ЗАЩИЩЁННЫХ ГРУПП (protected_groups.json) ==========
function Get-ProtectedEmails {
    param([string]$JsonPath)
    $map = @{}
    if (Test-Path $JsonPath) {
        try {
            $list = Get-Content -Path $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($entry in $list) {
                if ($entry.email) { $map[$entry.email.ToLowerInvariant().Trim()] = $entry }
            }
            Out-Log "[Information] Загружено $($map.Count) защищённых групп из $JsonPath."
        } catch {
            Out-Log "[Error] Ошибка чтения ${JsonPath}: $($_.Exception.Message)"
        }
    } else {
        Out-Log "[Warning] protected_groups.json не найден: $JsonPath"
    }
    return $map
}

# ========== УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ОТПРАВКИ ПИСЕМ ==========
function Send-Notification {
    param(
        [string]$MessageText = "",
        [string]$TableHtml = "",
        [array]$Attachments = @()
    )
    $MailTemplate = Get-Content -Path $TemplatePath -Raw -ErrorAction SilentlyContinue
    if (-not $MailTemplate) {
        $MailTemplate = "<html><body>{Message}<br>{Table}</body></html>"
    }
    $MailTemplate = $MailTemplate.Replace('{Message}', $MessageText).Replace('{Table}', $TableHtml)

    $MailMessage = @{
        From        = $MailSender
        To          = $MailRecipient
        Subject     = $MailSubject
        Body        = $MailTemplate
        SmtpServer  = $MailServer
        Port        = $MailPort
        Encoding    = "UTF8"
    }

    if ($Attachments.Count -gt 0) {
        $MailMessage.Attachments = $Attachments
    }

    try {
        Send-MailMessage @MailMessage -BodyAsHtml
        Out-Log "[Information] Notification email sent successfully."
    } catch {
        Out-Log "[Error] Failed to send notification email: $($_.Exception.Message)"
    }
}

# ========== ФУНКЦИЯ ИМПОРТА ДАННЫХ ИЗ EXCEL (ImportExcel) С ОЧИСТКОЙ ==========
function Import-DepartmentsFromExcel {
    param(
        [string]$FilePath,
        [string]$SheetName
    )
    try {
        if (-not (Test-Path $FilePath)) {
            throw "Файл '$FilePath' не найден."
        }

        $rawData = Import-Excel -Path $FilePath -WorksheetName $SheetName -ErrorAction Stop

        if ($rawData.Count -eq 0) {
            Out-Log "[Warning] Excel sheet '$SheetName' is empty or contains no data rows."
            return @()
        }

        $result = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($row in $rawData) {
            $id        = $row.'Уникальный идентификатор подразделения'
            $fullName  = $row.'Полное наименование пдоразделения'
            $headName  = $row.'Наименование подразделения (вершины)'
            $managerEN = $row.'Руководитель подразделения'

            if (-not $fullName) {
                $fullName = $row.PSObject.Properties | Where-Object { $_.Name -like '*полное*' } | Select-Object -First 1 -ExpandProperty Value
            }
            if (-not $id) {
                $id = $row.PSObject.Properties | Where-Object { $_.Name -like '*уникальный*' -or $_.Name -like '*идентификатор*' } | Select-Object -First 1 -ExpandProperty Value
            }
            if (-not $headName) {
                $headName = $row.PSObject.Properties | Where-Object { $_.Name -like '*вершины*' } | Select-Object -First 1 -ExpandProperty Value
            }
            if (-not $managerEN) {
                $managerEN = $row.PSObject.Properties | Where-Object { $_.Name -like '*руководитель*' } | Select-Object -First 1 -ExpandProperty Value
            }

            $fullName = if ($fullName) { $fullName.ToString().Trim() } else { '' }
            if ([string]::IsNullOrWhiteSpace($fullName)) { continue }

            # Срезаем ведущие нули у EmployeeNumber руководителя (Excel-формат "0000003966").
            # Исключение: руководитель с табельным номером "0000-00013" — единственный, у кого
            # номер содержит дефис и ведущие нули значимы (в AD он хранится как "0000-00013").
            # Для него оставляем значение как есть, иначе оно превратится в "-00013" и не совпадёт с AD.
            $managerEN = if ($managerEN) {
                $en = $managerEN.ToString().Trim()
                if ($en -eq '0000-00013') { $en } else { $en -replace '^0+', '' }
            } else { '' }

            $obj = [PSCustomObject]@{
                'Уникальный идентификатор подразделения' = if ($id) { $id.ToString().Trim() } else { '' }
                'Полное наименование пдоразделения'      = $fullName
                'Наименование подразделения (вершины)'   = if ($headName) { $headName.ToString().Trim() } else { '' }
                'Руководитель подразделения'             = $managerEN
            }
            $result.Add($obj)
        }

        # Фильтрация служебного мусора и дубликатов
        $filtered = $result | Where-Object {
            $_.'Полное наименование пдоразделения' -notmatch 'objName|полное наименование|наименование|^obj'
        }
        $unique = $filtered | Group-Object -Property 'Полное наименование пдоразделения' | ForEach-Object { $_.Group[0] }

        Out-Log "[Information] Imported $($unique.Count) unique departments from Excel (total raw rows: $($result.Count))."
        return $unique
    } catch {
        Out-Log "[Error] Excel import failed: $($_.Exception.Message)"
        throw
    }
}

# ========== ФУНКЦИЯ ПОЛУЧЕНИЯ ДАННЫХ ИЗ AD ==========
function Get-ADData {
    try {
        $OUs = @("OU=Users.Contractor,DC=transitcard,DC=ru", "OU=Users.Corp,DC=transitcard,DC=ru")

        $AllADUsers = [System.Collections.Generic.List[object]]::new()
        foreach ($ou in $OUs) {
            $filter = "Enabled -eq 'True' -and EmployeeNumber -like '*' -and Name -notlike '*[ADM]*' -and Name -notlike '*[VPN]*' -and Name -notlike '*[KZ]*'"
            $users = Get-ADUser -SearchBase $ou -Filter $filter -Properties EmployeeNumber, Department, Name, SamAccountName, Enabled, DistinguishedName
            foreach ($u in $users) { $AllADUsers.Add($u) }
        }
        Out-Log "[Information] Loaded $($AllADUsers.Count) AD users."

        $AllADGroups = Get-ADGroup -Filter "Description -eq 'Группа рассылки подразделения_v1.0'" -Properties DisplayName, Description, SamAccountName, CN, DistinguishedName, Members, MemberOf, mail
        Out-Log "[Information] Loaded $($AllADGroups.Count) target AD groups."

        $GroupByDisplayName = @{}
        $GroupByCN          = @{}
        $GroupByDistName    = @{}
        $GroupBySamName     = @{}
        foreach ($g in $AllADGroups) {
            if ($g.DisplayName)       { $GroupByDisplayName[$g.DisplayName] = $g }
            if ($g.CN)                { $GroupByCN[$g.CN] = $g }
            if ($g.DistinguishedName) { $GroupByDistName[$g.DistinguishedName] = $g }
            if ($g.SamAccountName)    { $GroupBySamName[$g.SamAccountName] = $g }
        }

        $UserBySam = @{}
        $UserByDN  = @{}
        $UserByEmpNum = @{}
        foreach ($u in $AllADUsers) {
            $UserBySam[$u.SamAccountName] = $u
            $UserByDN[$u.DistinguishedName] = $u
            if ($u.EmployeeNumber) {
                $UserByEmpNum[$u.EmployeeNumber.ToString().Trim()] = $u
            }
        }

        return @{
            Users                = $AllADUsers
            Groups               = $AllADGroups
            GroupByDisplayName   = $GroupByDisplayName
            GroupByCN            = $GroupByCN
            GroupByDistName      = $GroupByDistName
            GroupBySamName       = $GroupBySamName
            UserBySam            = $UserBySam
            UserByDN             = $UserByDN
            UserByEmpNum         = $UserByEmpNum
        }
    } catch {
        Out-Log "[Error] AD data retrieval failed: $($_.Exception.Message)"
        throw
    }
}

# ========== ФУНКЦИЯ ПОИСКА НЕКОРРЕКТНЫХ ГРУПП ==========
function Find-NonCorrADGroup {
    param(
        $AllNotes,
        $GroupByDisplayName
    )
    $NonCorrADGroup = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Department in $AllNotes) {
        $depName = $Department."Полное наименование пдоразделения"
        if ($depName -match '^obj') { continue }

        $foundGroup = Find-GroupMatch -SearchName $depName -Map $GroupByDisplayName
        if (-not $foundGroup) {
            $adGroup = Get-ADGroup -Filter "DisplayName -eq '$depName'" -Properties Description -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($adGroup) {
                $NonCorrADGroup.Add([PSCustomObject]@{
                    DisplayName = $depName
                    Description = "Группа рассылки подразделения_v1.0"
                    Reason      = "Некорректное поле Description"
                })
            } else {
                $NonCorrADGroup.Add([PSCustomObject]@{
                    DisplayName = $depName
                    Description = "Группа рассылки подразделения_v1.0"
                    Reason      = "Группа с таким DN не найдена"
                })
            }
        }
    }
    $NonCorrADGroup | Export-Csv -Path $Csv_NonCorrADGroup -Encoding UTF8 -NoTypeInformation
    return $NonCorrADGroup
}

# ========== ФУНКЦИЯ ПРОВЕРКИ ИЕРАРХИИ ==========
function Test-GroupHierarchy {
    param(
        $AllNotes,
        $GroupByDisplayName,
        $GroupByDistName
    )
    $NCMUpload = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Department in $AllNotes) {
        $CurrDep = $Department.'Полное наименование пдоразделения'
        $HeadDep = $Department.'Наименование подразделения (вершины)'

        # «Руководство» больше не верхнеуровневая группа: теперь она сама вложена
        # в кастомную «Все сотрудники». Проверяем именно это вхождение.
        if ($CurrDep -eq $LeadershipGroupName) {
            $HeadDep = $AllEmployeesGroupName
        }
        elseif ([string]::IsNullOrWhiteSpace($HeadDep)) {
            continue
        }
        # Везде, где 1С указывает вершиной «Руководство», фактическая вершина — «Все сотрудники».
        elseif ($HeadDep -eq $LeadershipGroupName) {
            $HeadDep = $AllEmployeesGroupName
        }

        $targetGroup = Find-GroupMatch -SearchName $CurrDep -Map $GroupByDisplayName
        $parentGroup = Find-GroupMatch -SearchName $HeadDep -Map $GroupByDisplayName
        if (-not $targetGroup -or -not $parentGroup) { continue }

        $isMember = $false
        foreach ($memberOfDN in $targetGroup.MemberOf) {
            $parentCandidate = $GroupByDistName[$memberOfDN]
            if ($parentCandidate -and $parentCandidate.Description -eq "Группа рассылки подразделения_v1.0") {
                if ($parentCandidate.DisplayName -eq $HeadDep) {
                    $isMember = $true
                    break
                }
            }
        }

        $groupsToExclude = @()
        if ($isMember) {
            foreach ($dn in $targetGroup.MemberOf) {
                $parent = $GroupByDistName[$dn]
                if ($parent -and $parent.Description -eq "Группа рассылки подразделения_v1.0") {
                    if ($parent.DisplayName -ne $HeadDep) {
                        $groupsToExclude += $parent.DisplayName
                    }
                }
            }
            if ($groupsToExclude.Count) {
                $NCMUpload.Add([PSCustomObject]@{
                    DGroup     = $CurrDep
                    WhereToAdd = $HeadDep
                    DeleteFrom = $groupsToExclude -join ','
                })
            }
        } else {
            $NCMUpload.Add([PSCustomObject]@{
                DGroup     = $CurrDep
                WhereToAdd = $HeadDep
                DeleteFrom = ''
            })
        }
    }
    $NCMUpload | Export-Csv -Path $Csv_NCMUpload -Encoding UTF8 -NoTypeInformation
    return $NCMUpload
}

# ========== ФУНКЦИЯ ПОИСКА НЕКОРРЕКТНЫХ ЧЛЕНОВ (ПОЛЬЗОВАТЕЛЕЙ) ==========
function Find-NonCorrUserMembership {
    param(
        $AllNotes,
        $GroupByDisplayName,
        $GroupByDistName,
        $UserByDN
    )

    # Карта: нормализованное имя группы -> EmployeeNumber руководителя (без ведущих нулей).
    # Один человек может руководить несколькими подразделениями — его DN будет в нескольких группах,
    # и в каждой из них он не должен попадать в NonCorrUserMembership, даже если AD.Department не совпадает.
    $managersByDept = @{}
    foreach ($d in $AllNotes) {
        $dn = $d.'Полное наименование пдоразделения'
        $en = $d.'Руководитель подразделения'
        if ($dn -and $en) {
            $managersByDept[$dn] = $en
        }
    }

    $NonCorrUserMembership = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Department in $AllNotes) {
        $depName = $Department.'Полное наименование пдоразделения'
        $group = Find-GroupMatch -SearchName $depName -Map $GroupByDisplayName
        if (-not $group) { continue }

        $managerEN = $managersByDept[$depName]

        foreach ($memberDN in $group.Members) {
            $memberGroup = $GroupByDistName[$memberDN]
            if ($memberGroup) { continue }
            $user = $UserByDN[$memberDN]
            if (-not $user) { continue }

            # Если пользователь — руководитель этого подразделения (по EmployeeNumber из 1С),
            # пропускаем, даже если его AD.Department отличается (кейс: один человек руководит двумя подразделениями).
            if ($managerEN -and $user.EmployeeNumber `
                -and ($user.EmployeeNumber.ToString().Trim() -eq $managerEN)) {
                continue
            }

            $userDep = if ($user.Department) { $user.Department.Trim() } else { '' }
            if ($userDep -ne $depName) {
                $NonCorrUserMembership.Add([PSCustomObject]@{
                    SamAccountName  = $user.SamAccountName
                    DeleteFromGroup = $depName
                })
            }
        }
    }
    $uniqueResult = $NonCorrUserMembership | Select-Object -Unique SamAccountName, DeleteFromGroup
    $uniqueResult | Export-Csv -Path $Csv_NonCorrUserMembership -Encoding UTF8 -NoTypeInformation
    return $uniqueResult
}

# ========== ФУНКЦИЯ ПОИСКА ПОЛЬЗОВАТЕЛЕЙ ДЛЯ ДОБАВЛЕНИЯ ==========
function Find-UsersToAdd {
    param(
        $AllADUsers,
        $GroupByCN
    )
    $UserDepToAdd = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($User in $AllADUsers) {
        $group = Find-GroupMatch -SearchName $User.Department -Map $GroupByCN
        if (-not $group) { continue }
        $isMember = $group.Members -contains $User.DistinguishedName
        if (-not $isMember) {
            $UserDepToAdd.Add([PSCustomObject]@{
                SamAccountName = $User.SamAccountName
                GroupToAdd     = $group.CN
            })
        }
    }
    $UserDepToAdd | Export-Csv -Path $Csv_UserDepToAdd -Encoding UTF8 -NoTypeInformation
    return $UserDepToAdd
}

# ========== ФУНКЦИЯ ПОИСКА НЕАКТИВНЫХ ПОЛЬЗОВАТЕЛЕЙ ДЛЯ УДАЛЕНИЯ ==========
function Find-UsersToExclude {
    param(
        $DGroups,
        $UserByDN
    )
    $UsersToExclude = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Group in $DGroups) {
        foreach ($memberDN in $Group.Members) {
            $user = $UserByDN[$memberDN]
            if (-not $user) { continue }
            if ($user.Enabled -ne $true) {
                $UsersToExclude.Add([PSCustomObject]@{
                    SamAccountName  = $user.SamAccountName
                    DeleteFromGroup = $Group.DisplayName
                })
            }
        }
    }
    $UsersToExclude | Export-Csv -Path $Csv_UsersToExclude -Encoding UTF8 -NoTypeInformation
    return $UsersToExclude
}

# ========== ФУНКЦИЯ ПОИСКА УСТАРЕВШИХ ГРУПП ДЛЯ УДАЛЕНИЯ ==========
function Find-ObsoleteGroups {
    param(
        $DGroups,
        $AllNotes,
        $ProtectedByEmail
    )
    $DGroupsToDel = [System.Collections.Generic.List[PSObject]]::new()

    # Актуальные имена подразделений из 1С (точное сопоставление по имени).
    $currentDepNames = @{}
    foreach ($dep in $AllNotes) {
        $n = $dep.'Полное наименование пдоразделения'
        if ($n) { $currentDepNames[$n.Trim()] = $true }
    }

    foreach ($Group in $DGroups) {
        $cn = if ($Group.CN)          { $Group.CN.Trim() }          else { '' }
        $dn = if ($Group.DisplayName) { $Group.DisplayName.Trim() } else { '' }

        # «Все сотрудники» — кастомная группа, которой намеренно нет в 1С. Не считаем устаревшей.
        if ($cn -eq $AllEmployeesGroupName -or $dn -eq $AllEmployeesGroupName) { continue }

        # Группа актуальна, если её имя (CN или DisplayName) есть среди подразделений 1С.
        if ($currentDepNames.ContainsKey($cn) -or $currentDepNames.ContainsKey($dn)) { continue }

        # Те же условия, что и на Шаге 4 скрипта синхронизации: удаляется только пустая
        # группа, которой нет в 1С и которая не защищена в protected_groups.json.
        $memberCount = if ($Group.Members) { $Group.Members.Count } else { 0 }
        $email       = if ($Group.mail) { $Group.mail.ToLowerInvariant().Trim() } else { $null }
        $isProtected = $email -and $ProtectedByEmail.ContainsKey($email)

        $reason =
            if ($memberCount -gt 0) {
                "Отсутствует в 1С, но в группе есть пользователи ($memberCount) — синхронизация не удаляет непустые группы"
            } elseif ($isProtected) {
                'Пустая и отсутствует в 1С, но защищена в protected_groups.json — не удаляется'
            } else {
                'Пустая и отсутствует в 1С — подлежит удалению'
            }

        $DGroupsToDel.Add([PSCustomObject]@{
            Mail        = $Group.mail
            DN          = $Group.DisplayName
            MemberCount = $memberCount
            Reason      = $reason
        })
    }
    $DGroupsToDel | Export-Csv -Path $Csv_DGroupsToDel -Encoding UTF8 -NoTypeInformation
    return $DGroupsToDel
}

# ========== ФУНКЦИЯ ФОРМИРОВАНИЯ HTML-ТАБЛИЦЫ СВОДКИ ==========
# Null-безопасный подсчёт записей.
# Нужен потому, что функции возвращают пустой набор как $null (пустой List при выводе
# разворачивается в 0 объектов), а @($null).Count = 1, а не 0. Measure-Object даёт
# корректно: $null/пусто -> 0, скаляр -> 1, массив -> N.
function Get-RecordCount {
    param($Data)
    return ($Data | Measure-Object).Count
}

function Get-SummaryHtmlTable {
    param(
        $NonCorrADGroup,
        $NCMUpload,
        $NonCorrUserMembership,
        $UserDepToAdd,
        $UsersToExclude,
        $DGroupsToDel
    )
    $rows = @(
        @{ Файл = "NonCorrADGroup.csv";         Описание = "Некорректно добавленные или несуществующие группы рассылки подразделений";          Записей = (Get-RecordCount $NonCorrADGroup) },
        @{ Файл = "NCMUpload.csv";              Описание = "Записи о некорректном вхождении группы в вышестоящий список рассылки";                Записей = (Get-RecordCount $NCMUpload) },
        @{ Файл = "NonCorrUserMembership.csv";  Описание = "Записи о некорректном вхождении пользователя в группу рассылки";                      Записей = (Get-RecordCount $NonCorrUserMembership) },
        @{ Файл = "UserDepToAdd.csv";           Описание = "Записи для включения пользователей в группы рассылки";                                Записей = (Get-RecordCount $UserDepToAdd) },
        @{ Файл = "UsersToExclude.csv";         Описание = "Записи о деактивированных пользователях, подлежащих исключению из группы";            Записей = (Get-RecordCount $UsersToExclude) },
        @{ Файл = "DGroupsToDel.csv";           Описание = "Список групп рассылки, подлежащих удалению";                                          Записей = (Get-RecordCount $DGroupsToDel) }
    )
    $html = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
    $html += "<tr><th>Файл</th><th>Описание</th><th>Количество записей</th></tr>"
    foreach ($row in $rows) {
        $html += "<tr><td>$($row.Файл)</td><td>$($row.Описание)</td><td style='text-align:center'>$($row.Записей)</td></tr>"
    }
    $html += "</table>"
    $html += "<br><b>Внимание!</b> Если добавлено новое подразделение и группа рассылки для него ранее не создавалась, необходимо придумать адрес группы рассылки и создать группу."
    return $html
}

# ========== ОСНОВНОЙ БЛОК СКРИПТА ==========
try {
    Out-Log "=== Начало выполнения ==="

    # 1. Импорт данных из Excel
    $AllNotes = Import-DepartmentsFromExcel -FilePath $ExcelPath -SheetName $ExcelSheet

    # 2. Загрузка данных из AD
    $ADData = Get-ADData
    $AllADUsers          = $ADData.Users
    $AllADGroups         = $ADData.Groups
    $GroupByDisplayName  = $ADData.GroupByDisplayName
    $GroupByCN           = $ADData.GroupByCN
    $GroupByDistName     = $ADData.GroupByDistName
    $UserByDN            = $ADData.UserByDN

    # 3. Некорректные группы рассылки
    $NonCorrADGroup = Find-NonCorrADGroup -AllNotes $AllNotes -GroupByDisplayName $GroupByDisplayName

    # 4. Проверка иерархии
    $NCMUpload = Test-GroupHierarchy -AllNotes $AllNotes -GroupByDisplayName $GroupByDisplayName -GroupByDistName $GroupByDistName

    # 5. Некорректные члены (пользователи) — с учётом руководителей нескольких подразделений
    $NonCorrUserMembership = Find-NonCorrUserMembership -AllNotes $AllNotes -GroupByDisplayName $GroupByDisplayName -GroupByDistName $GroupByDistName -UserByDN $UserByDN

    # 6. Пользователи для добавления
    $UserDepToAdd = Find-UsersToAdd -AllADUsers $AllADUsers -GroupByCN $GroupByCN

    # 7. Неактивные пользователи для исключения
    $DGroups = $AllADGroups
    $UsersToExclude = Find-UsersToExclude -DGroups $DGroups -UserByDN $UserByDN

    # 8. Устаревшие группы для удаления (с причиной по логике Шага 4 синхронизации)
    $protectedByEmail = Get-ProtectedEmails -JsonPath $ProtectedGroupsJsonPath
    $DGroupsToDel = Find-ObsoleteGroups -DGroups $DGroups -AllNotes $AllNotes -ProtectedByEmail $protectedByEmail

    # 9. Формирование сводной HTML-таблицы
    $SummaryTableHtml = Get-SummaryHtmlTable -NonCorrADGroup $NonCorrADGroup -NCMUpload $NCMUpload -NonCorrUserMembership $NonCorrUserMembership -UserDepToAdd $UserDepToAdd -UsersToExclude $UsersToExclude -DGroupsToDel $DGroupsToDel

    # 10. Подготовка вложений (только файлы, в которых есть данные)
    $Attachments = @()
    if ((Get-RecordCount $NonCorrADGroup) -gt 0)        { $Attachments += $Csv_NonCorrADGroup }
    if ((Get-RecordCount $NCMUpload) -gt 0)             { $Attachments += $Csv_NCMUpload }
    if ((Get-RecordCount $NonCorrUserMembership) -gt 0) { $Attachments += $Csv_NonCorrUserMembership }
    if ((Get-RecordCount $UserDepToAdd) -gt 0)          { $Attachments += $Csv_UserDepToAdd }
    if ((Get-RecordCount $UsersToExclude) -gt 0)        { $Attachments += $Csv_UsersToExclude }
    if ((Get-RecordCount $DGroupsToDel) -gt 0)          { $Attachments += $Csv_DGroupsToDel }

    # 11. Отправка итогового письма
    Send-Notification -TableHtml $SummaryTableHtml -Attachments $Attachments

    Out-Log "=== Успешное завершение ==="
} catch {
    Out-Log "[Critical] Script failed: $($_.Exception.Message)"
    exit 1
} finally {
    if ($mutex) {
        $mutex.ReleaseMutex()
    }
}
