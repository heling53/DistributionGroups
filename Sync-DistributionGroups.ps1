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
    Дополнительно ведутся файлы состояния:
      - protected_groups.json      — защита от удаления (ведётся вручную);
      - department_aliases.json    — старые названия подразделений, действующие на время переезда;
      - department_registry.json   — реестр групп: SID, адрес и вся история названий;
      - department_history.json    — журнал событий (создание, переименование, карантин, удаление).

    ПЕРЕИМЕНОВАНИЯ (шаг 0). 1С не переименовывает подразделение, а заводит новое и удаляет
    старое, поэтому его идентификатор переименование не переживает. Привязка берётся из AD:
    состав группы рассылки — это список людей, состоявших в подразделении на момент прошлой
    синхронизации. Когда подразделение исчезает из выгрузки, все уходы из него означают
    распад, и вопрос лишь в том, куда ушли люди. Совпавшая группа переименовывается на месте:
    SID, SamAccountName и почтовый адрес сохраняются, поэтому вложенность в группы доступа
    Jira/Confluence и любые ACL остаются нетронутыми.

    Порядок шагов важен: состав анализируется до того, как его меняет шаг 2.
.NOTES
    Все операции выполняются на одном DC ($DcServer), read-after-write согласован.
    Состав групп приводится к целевому разницей (добавить/убрать), без промежуточной очистки:
    очистка обнуляла бы доступы на время прогона, оставляла бы группы пустыми при обрыве
    прогона и стирала бы признак, по которому распознаются переименования.
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

    # Группы, которых синхронизация не касается вообще. «Все сотрудники» — вершина
    # дерева рассылки: в неё вложены «Руководство» и все отделы. В выгрузке 1С такого
    # узла нет и не будет — он остался со времён рукописного файла иерархии
    # ImportedDeps.csv, от которого отказались в пользу Подразделения.xlsx.
    # Группа ведётся вручную и не меняется, поэтому исключается на всех шагах:
    # ни состава, ни переименования, ни карантина, ни записей в файлы состояния.
    [string[]]$IgnoredGroups         = @('Все сотрудники'),

    # Подразделения, внутрь которых не вкладываются дочерние группы. «Руководство» —
    # вершина иерархии в выгрузке 1С, поэтому по справочнику в него попадали бы все
    # отделы. В AD дерево устроено иначе: и «Руководство», и все отделы вложены
    # напрямую в «Все сотрудники». Сама группа при этом обычная и управляется как все:
    # исключение отменяет только вложенность, но не состав из сотрудников.
    [string[]]$NoChildGroups         = @('Руководство'),

    # --- Файлы состояния ---
    [string]$AliasesJsonPath         = 'D:\Scripts\Группы рассылки\department_aliases.json',
    [string]$RegistryJsonPath        = 'D:\Scripts\Группы рассылки\department_registry.json',
    [string]$HistoryJsonPath         = 'D:\Scripts\Группы рассылки\department_history.json',
    [string]$RenameReportCsv         = 'D:\Scripts\Группы рассылки\RenameCandidates.csv',

    # --- Переименования ---
    # Старое название остаётся действующим столько дней: пока часть сотрудников
    # не сменила Department, они продолжают попадать в свою (переименованную) группу.
    [int]$AliasDays                  = 60,
    # Группа считается «молодой», если создана недавно: такая почти наверняка дубль,
    # порождённый уже идущим переименованием, а не самостоятельное подразделение.
    # Должно быть больше $QuarantineDays, иначе к моменту истечения карантина дубль
    # «повзрослеет» и перестанет распознаваться (принудительно выравнивается ниже).
    [int]$YoungGroupDays             = 45,
    # Порог автоматического переименования и обязательный отрыв от второго кандидата.
    [double]$AutoThreshold           = 0.80,
    [double]$AutoMargin              = 0.15,
    [double]$ManualThreshold         = 0.50,

    # --- Предохранители ---
    # Группа не опустошается автоматически: пустой целевой состав при непустом текущем —
    # типичный признак сбоя данных, а не роспуска подразделения.
    [int]$EmptyGuardMinMembers       = 3,
    # Доля исключений от общего числа членств, выше которой прогон считается аварийным
    # и состав не меняется вовсе.
    [double]$MassRemovalGuard        = 0.30,
    # Осиротевшая группа не удаляется сразу, а помечается и скрывается из адресной книги.
    # По истечении срока применяется лучший кандидат на переименование, а если его нет —
    # пустая группа удаляется.
    [int]$QuarantineDays             = 30
)

$ErrorActionPreference = 'Stop'

# Дубль под новым названием должен оставаться распознаваемым дольше, чем длится
# карантин: спорная пара применяется как раз в момент его истечения.
if ($YoungGroupDays -le $QuarantineDays) { $YoungGroupDays = $QuarantineDays + 15 }

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

$script:IgnoredKeys = @{}
foreach ($n in $IgnoredGroups) {
    if ($n -match '\S') { $script:IgnoredKeys[(Normalize-Name $n)] = $true }
}

$script:NoChildKeys = @{}
foreach ($n in $NoChildGroups) {
    if ($n -match '\S') { $script:NoChildKeys[(Normalize-Name $n)] = $true }
}

# Отсев исключённых групп. Применяется сразу после каждого чтения из AD, чтобы
# ни один шаг такую группу даже не увидел: так исключение нельзя случайно
# обойти, добавив новый шаг.
function Remove-IgnoredGroups {
    param($Groups)
    return @($Groups | Where-Object { -not $script:IgnoredKeys.ContainsKey((Normalize-Name $_.Name)) })
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
# Похожесть названий (для распознавания переименований)
# ==========================================
# Переименования обычно косметические («Отдел продаж» -> «Отдел продаж и маркетинга»,
# смена «отдел» на «управление», опечатка), поэтому похожесть названий здесь —
# полноценный сигнал, а не украшение.
$script:StopWords = @{}
foreach ($w in @('отдел','управление','департамент','служба','группа','сектор','дирекция',
                 'центр','бюро','подразделение','филиал','по','и','в','на','с','для')) {
    $script:StopWords[$w] = $true
}

# Токены без пунктуации, регистра и «ё», за вычетом слов-пустышек.
function Get-NameTokens {
    param([string]$Text)
    $t = (Normalize-Name $Text) -replace 'ё', 'е'
    $t = $t -replace '[^\p{L}\p{Nd}]+', ' '
    $all = @($t -split '\s+' | Where-Object { $_ })
    $significant = @($all | Where-Object { -not $script:StopWords.ContainsKey($_) })
    # Если после чистки не осталось ничего (название целиком из общих слов) — берём всё.
    if ($significant.Count -eq 0) { return $all }
    return $significant
}

function Get-LevenshteinDistance {
    param([string]$A, [string]$B)
    $n = $A.Length; $m = $B.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $prev = New-Object 'int[]' ($m + 1)
    $cur  = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($A[$i - 1] -eq $B[$j - 1]) { 0 } else { 1 }
            $cur[$j] = [Math]::Min([Math]::Min($cur[$j - 1] + 1, $prev[$j] + 1), $prev[$j - 1] + $cost)
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    return $prev[$m]
}

# Максимум из трёх мер, потому что переименования выглядят по-разному:
#   посимвольная  — опечатки и мелкие правки («Бухгалтерия» -> «Бухгалтерия»);
#   пословная     — перестановка и замена слов;
#   вложенность   — к названию приписали или из него убрали слово
#                   («Отдел продаж» -> «Отдел продаж и маркетинга»), самый частый случай.
# Вложенность занижена коэффициентом: одного её срабатывания не хватит,
# чтобы поднять пару до автоматического переименования без совпадения состава.
function Get-NameSimilarity {
    param([string]$A, [string]$B)
    $na = (Normalize-Name $A) -replace 'ё', 'е'
    $nb = (Normalize-Name $B) -replace 'ё', 'е'
    if (-not $na -or -not $nb) { return 0.0 }

    $maxLen = [Math]::Max($na.Length, $nb.Length)
    $lev    = 1.0 - ((Get-LevenshteinDistance -A $na -B $nb) / [double]$maxLen)

    $setA = @{}; foreach ($x in (Get-NameTokens $A)) { $setA[$x] = $true }
    $setB = @{}; foreach ($x in (Get-NameTokens $B)) { $setB[$x] = $true }
    $inter = 0
    foreach ($k in $setA.Keys) { if ($setB.ContainsKey($k)) { $inter++ } }
    $union = $setA.Count + $setB.Count - $inter
    $jac   = if ($union -gt 0) { $inter / [double]$union } else { 0.0 }

    $minCount = [Math]::Min($setA.Count, $setB.Count)
    $cont     = if ($minCount -gt 0) { 0.85 * ($inter / [double]$minCount) } else { 0.0 }

    return [Math]::Max([Math]::Max($lev, $jac), $cont)
}

# Разрешение цепочки псевдонимов: «Отдел А» -> «Отдел Б» -> «Отдел В».
# Ограничение по числу шагов защищает от кольца в данных.
function Resolve-DeptAlias {
    param([string]$Key, [hashtable]$AliasMap)
    $k = $Key
    for ($i = 0; $i -lt 5; $i++) {
        if (-not $AliasMap.ContainsKey($k)) { break }
        $next = $AliasMap[$k]
        if ($next -eq $k) { break }
        $k = $next
    }
    return $k
}

function Read-JsonArray {
    param([string]$Path)
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path $Path)) { return $out.ToArray() }
    try {
        foreach ($item in @(Get-Content -Path $Path -Raw | ConvertFrom-Json)) {
            if ($null -ne $item) { [void]$out.Add($item) }
        }
    } catch {
        Write-Log "Ошибка чтения ${Path}: $($_.Exception.Message)" 'ERROR'
    }
    return $out.ToArray()
}

function Write-JsonArray {
    param([string]$Path, $Items)
    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    # Обход через ArrayList намеренный. @() над System.Collections.Generic.List
    # роняет часть версий PowerShell с «Argument types do not match», а -InputObject
    # с массивом гарантирует, что и один элемент, и ноль дадут JSON-массив, а не скаляр.
    $out = New-Object System.Collections.ArrayList
    foreach ($i in $Items) { [void]$out.Add($i) }
    ConvertTo-Json -InputObject $out.ToArray() -Depth 6 | Set-Content -Path $Path -Encoding UTF8
}

$script:History = New-Object System.Collections.Generic.List[object]
function Add-HistoryEvent {
    param([string]$EventName, [string]$Group, [hashtable]$Details = @{})
    $entry = [ordered]@{
        Date  = (Get-Date).ToString('s')
        Event = $EventName
        Group = $Group
    }
    foreach ($k in $Details.Keys) { $entry[$k] = $Details[$k] }
    $script:History.Add([PSCustomObject]$entry)
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
    # Файлы состояния
    # ----------------------------------------
    $protectedByEmail = @{}
    foreach ($entry in (Read-JsonArray -Path $ProtectedGroupsJsonPath)) {
        if ($entry.email) { $protectedByEmail[$entry.email.ToLowerInvariant().Trim()] = $entry }
    }
    if ($protectedByEmail.Count -gt 0) {
        Write-Log "Загружено $($protectedByEmail.Count) защищённых групп." 'OK'
    } else {
        Write-Log "protected_groups.json не найден или пуст — защита от удаления отключена." 'WARN'
    }

    # Псевдоним — старое название, которое ещё встречается в атрибуте Department
    # у не переехавших сотрудников. Действует ограниченное время после переименования.
    $aliasList = New-Object System.Collections.Generic.List[object]
    $aliasMap  = @{}
    $today     = (Get-Date).Date
    foreach ($a in (Read-JsonArray -Path $AliasesJsonPath)) {
        if (-not $a.Old -or -not $a.New) { continue }
        if ($a.Expires -and ([datetime]$a.Expires) -lt $today) {
            Write-Log "Псевдоним истёк и снят: '$($a.Old)' -> '$($a.New)'" 'WARN'
            continue
        }
        $aliasList.Add($a)
        $aliasMap[(Normalize-Name $a.Old)] = (Normalize-Name $a.New)
    }
    Write-Log "Действующих псевдонимов подразделений: $($aliasList.Count)" 'OK'

    # Реестр — собственный источник постоянства: 1С его не даёт, поэтому история
    # названий каждой группы ведётся здесь, под ключом SamAccountName (он неизменен).
    $registry     = @{}
    $registryName = @{}   # нормализованное историческое название -> Sam
    foreach ($r in (Read-JsonArray -Path $RegistryJsonPath)) {
        if (-not $r.Sam) { continue }
        $registry[$r.Sam] = $r
        foreach ($n in @($r.Names)) {
            if ($n.Name) { $registryName[(Normalize-Name $n.Name)] = $r.Sam }
        }
    }
    Write-Log "Реестр групп: $($registry.Count) записей, $($registryName.Count) исторических названий" 'OK'

    # Именно перекладыванием в список: Read-JsonArray отдаёт результат массивом,
    # а у массива фиксированный размер и метода Add нет.
    foreach ($h in (Read-JsonArray -Path $HistoryJsonPath)) { $script:History.Add($h) }
    if ($script:History.Count -gt 0) { Write-Log "Журнал событий: $($script:History.Count) прежних записей" 'OK' }

    # ==========================================
    # Загрузка справочника 1С (Подразделения.xlsx) — единственный источник.
    # Отсюда: уникальные отделы, иерархия, руководители.
    # ==========================================
    Write-Log "Загрузка справочника 1С: $ExcelPath" 'STEP'
    if (-not (Test-Path $ExcelPath)) {
        throw "Справочник подразделений не найден: $ExcelPath"
    }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "Модуль ImportExcel не установлен. Установите: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel

    $excelHierarchy  = New-Object System.Collections.Generic.List[object]
    $excelManagers   = New-Object System.Collections.Generic.List[object]
    $departmentList  = New-Object System.Collections.Generic.List[string]
    $seenDept        = @{}   # нормализованное имя -> $true (дедупликация)
    $deptParentByKey = @{}   # нормализованное имя -> название вершины

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

        # Строки, ссылающиеся на исключённые группы, не берём: иначе шаг 2 будет
        # каждый прогон писать «родитель не найден» про группу, которой и не должно быть.
        if ($childName -and $script:IgnoredKeys.ContainsKey((Normalize-Name $childName))) { continue }

        if ($childName -and $parentName -and ($childName -ne $parentName)) {
            $parentKey = Normalize-Name $parentName
            # Вершину запоминаем всегда: она нужна как признак при распознавании
            # переименований. А вот вкладывать группы в исключённые и в помеченные
            # «без дочерних» не будем — там дерево собрано вручную по-своему.
            $deptParentByKey[(Normalize-Name $childName)] = $parentName
            if (-not $script:IgnoredKeys.ContainsKey($parentKey) -and
                -not $script:NoChildKeys.ContainsKey($parentKey)) {
                $excelHierarchy.Add([PSCustomObject]@{ Child = $childName; Parent = $parentName })
            }
        }
        if ($childName -and $managerEN) {
            $excelManagers.Add([PSCustomObject]@{ Department = $childName; EmployeeNumber = $managerEN })
        }

        # Уникальные подразделения: и сам узел (objFullName), и его родитель (objName),
        # чтобы для любого уровня иерархии существовала группа.
        foreach ($name in @($childName, $parentName)) {
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $norm = Normalize-Name $name
                if ($script:IgnoredKeys.ContainsKey($norm)) { continue }
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
    $excelGroupKeys = @{}
    $deptNameByKey  = @{}
    foreach ($d in $departments) {
        $k = Normalize-Name $d
        $excelGroupKeys[$k] = $true
        $deptNameByKey[$k]  = $d
    }

    # Руководитель по нормализованному названию подразделения.
    $managerKeyByDept = @{}
    foreach ($m in $excelManagers) { $managerKeyByDept[(Normalize-Name $m.Department)] = $m.EmployeeNumber }

    # ==========================================
    # Загрузка пользователей AD
    # ==========================================
    Write-Log "Загрузка пользователей AD" 'STEP'

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

    $users = @(Get-ADUser -Filter $filterUsers `
                          -Properties Department, EmployeeNumber, Name `
                          -ResultPageSize 1000 -Server $DcServer)

    # Исключение: сотрудники $UsersCorpOU без табельного номера. Основной фильтр
    # требует EmployeeNumber, и без этого блока они не попали бы ни в одну группу.
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
        $exceptionUsers = @(Get-ADUser -Filter $filterNoEmpNum `
                                       -SearchBase $UsersCorpOU `
                                       -Properties Department, Name `
                                       -ResultPageSize 1000 -Server $DcServer)
        Write-Log "Исключений (без EmployeeNumber в $UsersCorpOU): $($exceptionUsers.Count)" 'OK'
    } catch {
        Write-Log "Не удалось получить пользователей из ${UsersCorpOU}: $($_.Exception.Message)" 'ERROR'
    }

    $allUsers = @($users + $exceptionUsers)

    # Раскладка пользователей по группам с учётом псевдонимов: пока сотрудник
    # не переехал, его старое Department ведёт в ту же (уже переименованную) группу.
    function Build-UsersByDept {
        param($UserList, [hashtable]$AliasMap)
        $map = @{}
        foreach ($u in $UserList) {
            if ($u.Department -notmatch '\S') { continue }
            $key = Resolve-DeptAlias -Key (Normalize-Name $u.Department) -AliasMap $AliasMap
            if (-not $map.ContainsKey($key)) {
                $map[$key] = New-Object System.Collections.Generic.List[object]
            }
            $map[$key].Add($u)
        }
        return $map
    }

    $usersByDept   = Build-UsersByDept -UserList $allUsers -AliasMap $aliasMap
    $usersByEmpNum = @{}
    $deptByPerson  = @{}   # ключ человека -> фактическое подразделение (без псевдонимов)
    $personKeyByDN = @{}
    $peopleInDept  = @{}   # фактическое подразделение -> число людей

    foreach ($u in $allUsers) {
        if ($u.Department -notmatch '\S') { continue }
        $rawKey = Normalize-Name $u.Department

        # Ключ человека: табельный номер, если он есть, иначе SamAccountName.
        # Табельный переживает и смену ФИО, и смену подразделения.
        $personKey = if ($u.EmployeeNumber) { 'EN:' + $u.EmployeeNumber.ToString().Trim() }
                     else { 'SAM:' + $u.SamAccountName }
        $deptByPerson[$personKey] = $rawKey
        $personKeyByDN[$u.DistinguishedName] = $personKey
        if (-not $peopleInDept.ContainsKey($rawKey)) { $peopleInDept[$rawKey] = 0 }
        $peopleInDept[$rawKey]++

        if ($u.EmployeeNumber) { $usersByEmpNum[$u.EmployeeNumber.ToString().Trim()] = $u }
    }
    Write-Log "AD: пользователей учтено $($personKeyByDN.Count)" 'OK'

    # ==========================================
    # ШАГ 0. Распознавание переименований
    # ==========================================
    # Выполняется до любых изменений состава: состав групп на этот момент —
    # снимок подразделений на прошлую синхронизацию, и это единственная опора.
    Write-Log "Шаг 0: распознавание переименований подразделений" 'STEP'

    $adGroups = Remove-IgnoredGroups (Get-ADGroup -LDAPFilter "(description=$GroupDesc)" -SearchBase $TargetOU -Server $DcServer `
                            -Properties Name, DisplayName, SamAccountName, mail, Member, MemberOf, `
                                        ManagedBy, whenCreated, info, objectSid)

    $groupNameByDN = @{}
    $managedDNs    = @{}
    foreach ($g in $adGroups) {
        $groupNameByDN[$g.DistinguishedName] = $g.Name
        $managedDNs[$g.DistinguishedName]    = $true
    }

    $groupInfo  = New-Object System.Collections.Generic.List[object]
    $groupByKey = @{}
    foreach ($g in $adGroups) {
        # Состав только из людей: вложенные группы — это иерархия, а не сотрудники.
        $memberKeys = @{}
        foreach ($dn in $g.Member) {
            if ($managedDNs.ContainsKey($dn))     { continue }
            if ($personKeyByDN.ContainsKey($dn))  { $memberKeys[$personKeyByDN[$dn]] = $true }
        }

        # Родитель в иерархии: группа-владелец с тем же служебным Description.
        $parentName = ''
        foreach ($dn in $g.MemberOf) {
            if ($managedDNs.ContainsKey($dn)) { $parentName = $groupNameByDN[$dn]; break }
        }

        # Руководитель группы: ManagedBy проставляется этим же скриптом из 1С,
        # то есть хранит руководителя подразделения на момент до переименования.
        $managerEN = ''
        if ($g.ManagedBy -and $personKeyByDN.ContainsKey($g.ManagedBy)) {
            $mk = $personKeyByDN[$g.ManagedBy]
            if ($mk -like 'EN:*') { $managerEN = $mk.Substring(3) }
        }

        $quarantineSince = $null
        if ($g.info -and ($g.info -match 'QUARANTINE:(\d{4}-\d{2}-\d{2})')) {
            $quarantineSince = [datetime]$matches[1]
        }

        # Все известные названия группы, включая прежние: если 1С заведёт отдел
        # под названием, которое эта группа уже носила, привязка будет точной.
        $knownNames = New-Object System.Collections.Generic.List[string]
        $knownNames.Add($g.Name)
        if ($registry.ContainsKey($g.SamAccountName)) {
            foreach ($n in @($registry[$g.SamAccountName].Names)) {
                if ($n.Name -and ($knownNames -notcontains $n.Name)) { $knownNames.Add($n.Name) }
            }
        }

        $obj = [PSCustomObject]@{
            Name            = $g.Name
            Key             = Normalize-Name $g.Name
            Dn              = $g.DistinguishedName
            Sam             = $g.SamAccountName
            Sid             = if ($g.objectSid) { $g.objectSid.Value } else { '' }
            Mail            = $g.mail
            MemberKeys      = $memberKeys
            ParentName      = $parentName
            ManagerEN       = $managerEN
            QuarantineSince = $quarantineSince
            KnownNames      = $knownNames
            ExternalParents = @($g.MemberOf | Where-Object { -not $managedDNs.ContainsKey($_) })
            AgeDays         = if ($g.whenCreated) { ((Get-Date) - $g.whenCreated).TotalDays } else { [double]::MaxValue }
        }
        $groupInfo.Add($obj)
        $groupByKey[$obj.Key] = $obj
    }

    # Защищённые группы исключаются из механики переименований целиком: администратор
    # закрепил их вручную, и ни переименовать, ни отставить в сторону их нельзя.
    function Test-Protected {
        param($GroupObj)
        if (-not $GroupObj -or -not $GroupObj.Mail) { return $false }
        return $protectedByEmail.ContainsKey($GroupObj.Mail.ToLowerInvariant().Trim())
    }

    # Осиротевшая группа — та, чьего названия больше нет в выгрузке 1С. Именно это
    # событие отличает переименование от перевода сотрудника: при переводе исходное
    # подразделение остаётся в справочнике, и уход одного человека ничего не значит.
    $orphans = @($groupInfo | Where-Object {
        (-not $excelGroupKeys.ContainsKey($_.Key)) -and -not (Test-Protected $_)
    })

    # Получателем может быть подразделение, у которого группы ещё нет, либо у которого
    # группа создана недавно (дубль от уже идущего переименования). Подразделение
    # с давно существующей группой получателем не бывает: туда переводят по-настоящему.
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($k in $excelGroupKeys.Keys) {
        $existing = $groupByKey[$k]
        if ($null -eq $existing) {
            $targets.Add([PSCustomObject]@{ Key = $k; Duplicate = $null })
        } elseif ($existing.AgeDays -le $YoungGroupDays -and -not (Test-Protected $existing)) {
            $targets.Add([PSCustomObject]@{ Key = $k; Duplicate = $existing })
        }
    }
    Write-Log ("Осиротевших групп: {0}; подразделений-получателей: {1}" -f $orphans.Count, $targets.Count) 'OK'

    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($o in $orphans) {
        # Ушедшие: те, кто числился в группе, а сейчас имеет другое подразделение.
        # Уволенных в выборке AD нет, и в расчёт они не идут.
        $movers = New-Object System.Collections.Generic.List[string]
        foreach ($k in $o.MemberKeys.Keys) {
            if (-not $deptByPerson.ContainsKey($k)) { continue }
            if ($deptByPerson[$k] -ne $o.Key)       { $movers.Add($k) }
        }

        foreach ($n in $targets) {
            $movedHere = 0
            foreach ($k in $movers) { if ($deptByPerson[$k] -eq $n.Key) { $movedHere++ } }
            $newCount = if ($peopleInDept.ContainsKey($n.Key)) { $peopleInDept[$n.Key] } else { 0 }
            $newName  = $deptNameByKey[$n.Key]

            # Точное совпадение с одним из прежних названий этой же группы — привязка
            # без домыслов: отдел вернулся к названию, которое группа уже носила.
            $historyHit = ($o.KnownNames | Where-Object { (Normalize-Name $_) -eq $n.Key }).Count -gt 0

            # Два вопроса, и оба должны получить ответ «да»:
            #   доля ушедших, попавших именно сюда  — иначе это разделение отдела;
            #   доля нового отдела, пришедшая оттуда — иначе это слияние или перевод.
            # Берём минимум: он не даёт одному высокому показателю замаскировать другой.
            # Так работает и постепенный переезд: четверо ушедших из двенадцати, все
            # четверо здесь, и в новом отделе пока только они — обе доли равны единице.
            $hasMembers  = ($movers.Count -gt 0 -and $newCount -gt 0)
            $memberScore = 0.0
            if ($hasMembers) {
                $movedShare  = $movedHere / [double]$movers.Count
                $purity      = $movedHere / [double]$newCount
                $memberScore = [Math]::Min($movedShare, $purity)
            }

            # Похожесть считаем по всем известным названиям группы и берём лучшую.
            $nameScore = 0.0
            foreach ($kn in $o.KnownNames) {
                $s = Get-NameSimilarity -A $kn -B $newName
                if ($s -gt $nameScore) { $nameScore = $s }
            }

            $newManagerEN = $managerKeyByDept[$n.Key]
            $hasManager   = ($o.ManagerEN -and $newManagerEN)
            $managerScore = if ($hasManager -and ($o.ManagerEN -eq $newManagerEN)) { 1.0 } else { 0.0 }

            $newParent   = $deptParentByKey[$n.Key]
            $hasParent   = ($o.ParentName -and $newParent)
            $parentScore = if ($hasParent -and ((Normalize-Name $o.ParentName) -eq (Normalize-Name $newParent))) { 1.0 } else { 0.0 }

            # Веса нормируются по доступным сигналам: пока никто не переехал, состава
            # нет, и оценку несут название, руководитель и вершина.
            $sum = 0.0; $wsum = 0.0
            if ($hasMembers) { $sum += 0.45 * $memberScore;  $wsum += 0.45 }
            $sum += 0.30 * $nameScore; $wsum += 0.30
            if ($hasManager) { $sum += 0.15 * $managerScore; $wsum += 0.15 }
            if ($hasParent)  { $sum += 0.10 * $parentScore;  $wsum += 0.10 }
            $score = if ($wsum -gt 0) { $sum / $wsum } else { 0.0 }
            if ($historyHit) { $score = 1.0 }

            if ($score -lt 0.30) { continue }   # заведомо чужие пары в отчёт не тащим

            $pairs.Add([PSCustomObject]@{
                Old = $o; NewKey = $n.Key; NewName = $newName; Duplicate = $n.Duplicate
                MoversCount = $movers.Count; MovedHere = $movedHere; NewCount = $newCount
                HasMembers  = $hasMembers
                HistoryHit  = $historyHit
                MemberScore = [Math]::Round($memberScore, 3)
                NameScore   = [Math]::Round($nameScore, 3)
                ManagerHit  = [int]$managerScore
                ParentHit   = [int]$parentScore
                Score       = [Math]::Round($score, 3)
            })
        }
    }

    # Жадное сопоставление один к одному: одна группа не может быть переименована
    # дважды, и одно подразделение не может забрать две старые группы — иначе
    # слияние отделов молча уничтожит одну из них вместе с её доступами.
    $sortedPairs  = @($pairs | Sort-Object -Property Score -Descending)
    $usedOld      = @{}
    $usedNew      = @{}
    $renameReport = New-Object System.Collections.Generic.List[object]

    foreach ($p in $sortedPairs) {
        if ($usedOld.ContainsKey($p.Old.Key) -or $usedNew.ContainsKey($p.NewKey)) { continue }

        # Отрыв от лучшей альтернативы по любой из двух сторон пары. Без него
        # разделение отдела надвое молча свелось бы к случайному из двух вариантов.
        $rival = $sortedPairs | Where-Object {
            ($_.Old.Key -eq $p.Old.Key -or $_.NewKey -eq $p.NewKey) -and
            -not ($_.Old.Key -eq $p.Old.Key -and $_.NewKey -eq $p.NewKey)
        } | Select-Object -First 1
        $margin = if ($rival) { $p.Score - $rival.Score } else { $p.Score }

        # У подразделения-получателя уже есть своя группа: чтобы отдать ему имя,
        # придётся отставить её в сторону. Это разрушительно — у той группы свой SID
        # и свои выданные доступы, — поэтому планка выше и состав обязан почти
        # полностью совпадать. Так распад отдела, чьи люди разошлись по соседнему
        # подразделению, не сможет отобрать имя у живой группы.
        $displaces      = ($null -ne $p.Duplicate)
        $memberBar      = if ($displaces) { 0.8 } else { 0.5 }

        # Автомат допускается только при подтверждённом переездом составе: похожесть
        # названий сама по себе может совпасть у двух родственных отделов.
        # Исключение — точное попадание в прежнее название этой же группы.
        $decision =
            if ($p.HistoryHit -and -not $displaces)        { 'Auto' }
            elseif ($p.Score -ge $AutoThreshold -and $margin -ge $AutoMargin -and
                    $p.HasMembers -and $p.MemberScore -ge $memberBar) { 'Auto' }
            elseif ($p.Score -ge $ManualThreshold)         { 'Manual' }
            else                                           { 'Weak' }

        # Спорную пару, провисевшую в карантине дольше срока, применяем сами: держать
        # группу скрытой бесконечно хуже, чем принять лучший вариант. Но только если
        # имя свободно — вытеснять чужую живую группу по таймауту нельзя никогда.
        $reason = if ($p.HistoryHit) { 'Прежнее название группы' } else { 'Оценка' }
        if ($decision -eq 'Manual' -and -not $displaces -and $p.Old.QuarantineSince) {
            $daysWaiting = ((Get-Date) - $p.Old.QuarantineSince).TotalDays
            if ($daysWaiting -ge $QuarantineDays) {
                $decision = 'Auto'
                $reason   = "Карантин истёк ($([int]$daysWaiting) дн.), применён лучший кандидат"
            }
        }

        $usedOld[$p.Old.Key] = $true
        $usedNew[$p.NewKey]  = $true

        $renameReport.Add([PSCustomObject]@{
            Decision      = $decision
            Reason        = $reason
            OldGroup      = $p.Old.Name
            NewDepartment = $p.NewName
            Score         = $p.Score
            Margin        = [Math]::Round($margin, 3)
            MemberScore   = $p.MemberScore
            Moved         = "$($p.MovedHere)/$($p.MoversCount) ушедших -> отдел из $($p.NewCount)"
            NameScore     = $p.NameScore
            ManagerHit    = $p.ManagerHit
            ParentHit     = $p.ParentHit
            DuplicateName = if ($p.Duplicate) { $p.Duplicate.Name } else { '' }
            Sam           = $p.Old.Sam
            Mail          = $p.Old.Mail
            Dn            = $p.Old.Dn
        })
    }

    # Осиротевшие группы без кандидата: они уйдут в карантин на шаге 4.
    foreach ($o in $orphans) {
        if ($usedOld.ContainsKey($o.Key)) { continue }
        $renameReport.Add([PSCustomObject]@{
            Decision = 'NoMatch'; Reason = 'Кандидатов нет'; OldGroup = $o.Name; NewDepartment = ''
            Score = 0; Margin = 0; MemberScore = 0
            Moved = "состав $($o.MemberKeys.Count)"
            NameScore = 0; ManagerHit = 0; ParentHit = 0; DuplicateName = ''
            Sam = $o.Sam; Mail = $o.Mail; Dn = $o.Dn
        })
    }

    if ($renameReport.Count -gt 0) {
        $reportDir = Split-Path $RenameReportCsv -Parent
        if ($reportDir -and -not (Test-Path $reportDir)) { New-Item -ItemType Directory -Path $reportDir -Force | Out-Null }
        $renameReport | Export-Csv -Path $RenameReportCsv -Encoding UTF8 -NoTypeInformation
        foreach ($grp in ($renameReport | Group-Object Decision)) {
            Write-Log ("Кандидаты на переименование, {0}: {1}" -f $grp.Name, $grp.Count) 'OK'
        }
    }

    # ----------------------------------------
    # Применение переименований
    # ----------------------------------------
    $renamedKeys = @{}
    foreach ($r in ($renameReport | Where-Object { $_.Decision -eq 'Auto' })) {
        $survivor = $groupByKey[(Normalize-Name $r.OldGroup)]

        # Дубль под новым названием создан прошлым прогоном. Освобождаем имя,
        # отставляя дубль в сторону; удалять его будет шаг 4 на общих основаниях.
        $inheritedParents = @()
        if ($r.DuplicateName) {
            $dupGroup = $groupByKey[(Normalize-Name $r.DuplicateName)]
            $parkName = "ZZ-Дубль-{0}-{1:yyyyMMdd}" -f $r.DuplicateName, (Get-Date)
            if ($parkName.Length -gt 64) { $parkName = $parkName.Substring(0, 64) }
            if ($PSCmdlet.ShouldProcess($r.DuplicateName, "Rename-ADObject -> $parkName")) {
                try {
                    Rename-ADObject -Identity $dupGroup.Dn -NewName $parkName -Server $DcServer
                    # Доступ, выданный дублю за дни его существования (вложение в группы
                    # Jira/Confluence и прочие), переносим на выживающую группу — иначе
                    # он исчезнет вместе с дублем.
                    $inheritedParents = @($dupGroup.ExternalParents)
                    Add-HistoryEvent -EventName 'DuplicateParked' -Group $r.DuplicateName -Details @{
                        ParkedAs = $parkName; Survivor = $r.Sam; InheritedParents = $inheritedParents
                    }
                    Write-Log "Дубль отставлен: '$($r.DuplicateName)' -> '$parkName'" 'WARN'
                } catch {
                    Write-Log "Не удалось освободить имя '$($r.DuplicateName)': $($_.Exception.Message)" 'ERROR'
                    continue
                }
            }
        }

        # Переименование на месте: SID, SamAccountName и почтовый адрес прежние,
        # поэтому вложенность в группы доступа и ACL остаются нетронутыми.
        # CN меняется намеренно — дальше группы ищутся по Name, и иначе возник бы дубль.
        if ($PSCmdlet.ShouldProcess("$($r.OldGroup) -> $($r.NewDepartment)", "Rename-ADObject + DisplayName")) {
            try {
                Rename-ADObject -Identity $r.Dn -NewName $r.NewDepartment -Server $DcServer
                Set-ADGroup -Identity $r.Sam -DisplayName $r.NewDepartment -Server $DcServer
                # Карантинную пометку снимаем: группа снова соответствует 1С.
                try {
                    Set-ADGroup -Identity $r.Sam -Clear info -Server $DcServer
                    Set-DistributionGroup -Identity $r.Sam -HiddenFromAddressListsEnabled $false -DomainController $DcServer
                } catch { }

                foreach ($parentDn in $inheritedParents) {
                    try {
                        Add-ADGroupMember -Identity $parentDn -Members (Get-ADGroup -Identity $r.Sam -Server $DcServer) -Server $DcServer
                        Write-Log "Доступ дубля перенесён: $($r.Sam) -> $parentDn" 'OK'
                    } catch {
                        if ($_.Exception.Message -notmatch 'already a member') {
                            Write-Log "Не удалось перенести вхождение в ${parentDn}: $($_.Exception.Message)" 'ERROR'
                        }
                    }
                }

                $renamedKeys[(Normalize-Name $r.OldGroup)] = Normalize-Name $r.NewDepartment
                $aliasList.Add([PSCustomObject]@{
                    Old     = $r.OldGroup
                    New     = $r.NewDepartment
                    Since   = (Get-Date).ToString('yyyy-MM-dd')
                    Expires = (Get-Date).AddDays($AliasDays).ToString('yyyy-MM-dd')
                })
                Add-HistoryEvent -EventName 'Rename' -Group $r.NewDepartment -Details @{
                    OldName = $r.OldGroup; Sam = $r.Sam; Sid = $survivor.Sid; Mail = $r.Mail
                    Score = $r.Score; Margin = $r.Margin; Moved = $r.Moved; Reason = $r.Reason
                }
                Write-Log "Переименована: '$($r.OldGroup)' -> '$($r.NewDepartment)' ($($r.Reason), SID сохранён, $($r.Sam))" 'OK'
            } catch {
                Write-Log "Ошибка переименования '$($r.OldGroup)': $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # Псевдонимы для только что переименованных: сотрудники, ещё не переехавшие
    # на новое Department, продолжают попадать в свою (уже переименованную) группу.
    # Без этого при постепенном переезде часть отдела выпала бы из группы.
    if ($renamedKeys.Count -gt 0) {
        foreach ($oldKey in $renamedKeys.Keys) { $aliasMap[$oldKey] = $renamedKeys[$oldKey] }
        $usersByDept = Build-UsersByDept -UserList $allUsers -AliasMap $aliasMap
    }
    Write-JsonArray -Path $AliasesJsonPath -Items $aliasList
    Write-Log "Псевдонимов действует: $($aliasList.Count)" 'OK'

    $pendingRenames = @($renameReport | Where-Object { $_.Decision -eq 'Manual' })
    if ($pendingRenames.Count -gt 0) {
        Write-Log "Требуют решения человека ($($pendingRenames.Count)), см. $RenameReportCsv" 'WARN'
        foreach ($r in $pendingRenames) {
            Write-Log "    '$($r.OldGroup)' -> '$($r.NewDepartment)' (оценка $($r.Score), $($r.Moved))" 'WARN'
        }
    }

    # ==========================================
    # ШАГ 1. Создание недостающих групп
    # ==========================================
    Write-Log "Шаг 1: создание групп из справочника 1С" 'STEP'

    $existingGroups = Remove-IgnoredGroups (Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                  -SearchBase $TargetOU -Server $DcServer `
                                  -Properties Name, SamAccountName, mail, DisplayName, Member)
    $groupsByName = @{}
    $usedSam      = @{}
    foreach ($g in $existingGroups) {
        $groupsByName[(Normalize-Name $g.Name)] = $g
        $usedSam[$g.SamAccountName] = $true
    }

    foreach ($dept in $departments) {
        $groupName = $dept
        if ($groupsByName.ContainsKey((Normalize-Name $groupName))) { continue }

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

                Add-HistoryEvent -EventName 'Create' -Group $groupName -Details @{ Sam = $sam }
                Write-Log "Создана группа: $groupName ($sam)" 'OK'
            } catch {
                Write-Log "Ошибка создания '$groupName': $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 2. Приведение состава групп к целевому (разницей)
    # ==========================================
    # Состав не очищается и не наливается заново: сначала считается целевой набор,
    # затем применяется только разница. Полная очистка обнуляла бы доступы на время
    # прогона и оставляла бы группы пустыми, если прогон прервался на середине.
    Write-Log "Шаг 2: приведение состава групп к целевому" 'STEP'

    $managedGroups = Remove-IgnoredGroups (Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                 -SearchBase $TargetOU -Server $DcServer `
                                 -Properties Member, mail, Name, SamAccountName, info, objectSid)
    $managedGroupsMap = @{}
    $managedGroupDNs  = @{}
    foreach ($g in $managedGroups) {
        $managedGroupsMap[(Normalize-Name $g.Name)] = $g
        $managedGroupDNs[$g.DistinguishedName]      = $true
    }

    # Дочерние группы по родителю: вложенность из 1С — такие же члены группы,
    # и учитывать их надо здесь же, иначе разница вычтет их как лишние.
    $childrenByParent = @{}
    foreach ($row in $excelHierarchy) {
        $parentKey = Normalize-Name $row.Parent
        $childKey  = Normalize-Name $row.Child
        if (-not $managedGroupsMap.ContainsKey($parentKey)) { Write-Log "Родитель '$($row.Parent)' не найден" 'WARN'; continue }
        if (-not $managedGroupsMap.ContainsKey($childKey))  { Write-Log "Потомок '$($row.Child)' не найден"  'WARN'; continue }
        if (-not $childrenByParent.ContainsKey($parentKey)) {
            $childrenByParent[$parentKey] = New-Object System.Collections.Generic.List[string]
        }
        $childrenByParent[$parentKey].Add($managedGroupsMap[$childKey].DistinguishedName)
    }

    # --- Планирование: считаем все изменения, ничего не применяя ---
    $plan          = New-Object System.Collections.Generic.List[object]
    $totalCurrent  = 0
    $totalToRemove = 0

    foreach ($g in $managedGroups) {
        $key = Normalize-Name $g.Name

        # Состав осиротевшей группы замораживаем целиком. Причин две. Во-первых, это
        # улика: именно по нему на следующих прогонах распознаётся переименование,
        # и, опустошив группу, мы ослепим себя. Во-вторых, отбирать доступ в тот же
        # день, когда подразделение исчезло из 1С, рано — люди ещё переезжают.
        # Размораживается на шаге 4 по истечении карантина.
        if (-not $excelGroupKeys.ContainsKey($key)) {
            Write-Log "Состав заморожен (нет в 1С, идёт карантин): $($g.Name)" 'WARN'
            continue
        }

        $desired = @{}
        if ($usersByDept.ContainsKey($key)) {
            foreach ($u in $usersByDept[$key]) { $desired[$u.DistinguishedName] = $true }
        }
        $desiredUserCount = $desired.Count

        # Руководитель подразделения из 1С: он может числиться в другом отделе,
        # но в свою группу входить обязан.
        if ($managerKeyByDept.ContainsKey($key)) {
            $en = $managerKeyByDept[$key]
            if ($usersByEmpNum.ContainsKey($en)) {
                $desired[$usersByEmpNum[$en].DistinguishedName] = $true
            } else {
                Write-Log "Руководитель EmployeeNumber=$en для '$($g.Name)' не найден в AD" 'WARN'
            }
        }
        if ($childrenByParent.ContainsKey($key)) {
            foreach ($dn in $childrenByParent[$key]) { $desired[$dn] = $true }
        }

        $current = @{}
        foreach ($dn in $g.Member) { $current[$dn] = $true }

        $toAdd = @($desired.Keys | Where-Object { -not $current.ContainsKey($_) })
        $toDel = @($current.Keys | Where-Object { -not $desired.ContainsKey($_) })

        $totalCurrent  += $current.Count
        $totalToRemove += $toDel.Count

        # Предохранитель. Сюда доходят только подразделения, которые есть в 1С,
        # то есть живые. Пустой целевой состав у живого подразделения — признак
        # сбоя данных (не прочиталась выгрузка, слетел Department у сотрудников),
        # а не роспуска. Такую группу не трогаем.
        # Считаем именно людей: вложенные группы к признаку сбоя отношения не имеют,
        # иначе вершина иерархии без своих сотрудников замерла бы навсегда.
        $currentUserCount = @($current.Keys | Where-Object { -not $managedGroupDNs.ContainsKey($_) }).Count
        if ($desiredUserCount -eq 0 -and $currentUserCount -ge $EmptyGuardMinMembers) {
            Write-Log "Пропуск '$($g.Name)': целевой состав пуст при $currentUserCount текущих — проверьте данные" 'WARN'
            continue
        }

        if ($toAdd.Count -gt 0 -or $toDel.Count -gt 0) {
            $plan.Add([PSCustomObject]@{ Group = $g; ToAdd = $toAdd; ToDel = $toDel })
        }
    }

    # Общий предохранитель: массовое исключение по всем группам сразу означает
    # испорченные входные данные, а не реорганизацию. Состав не меняем вовсе.
    $removalShare = if ($totalCurrent -gt 0) { $totalToRemove / [double]$totalCurrent } else { 0 }
    if ($removalShare -gt $MassRemovalGuard) {
        Write-Log ("АВАРИЯ: под исключение попало {0} из {1} членств ({2:P0}) при пороге {3:P0}. Состав не изменён." -f `
            $totalToRemove, $totalCurrent, $removalShare, $MassRemovalGuard) 'ERROR'
        Add-HistoryEvent -EventName 'MassRemovalBlocked' -Group '' -Details @{
            ToRemove = $totalToRemove; Total = $totalCurrent
        }
    } else {
        foreach ($item in $plan) {
            $g = $item.Group
            if ($item.ToAdd.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess($g.Name, "Add $($item.ToAdd.Count) members")) {
                    try {
                        Add-ADGroupMember -Identity $g -Members $item.ToAdd -Server $DcServer
                        Write-Log "+$($item.ToAdd.Count) -> $($g.Name)" 'OK'
                    } catch { Write-Log "Add fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
                }
            }
            if ($item.ToDel.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess($g.Name, "Remove $($item.ToDel.Count) members")) {
                    try {
                        Remove-ADGroupMember -Identity $g -Members $item.ToDel -Server $DcServer -Confirm:$false
                        Write-Log "-$($item.ToDel.Count) <- $($g.Name)" 'OK'
                    } catch { Write-Log "Remove fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
                }
            }
        }
        Write-Log ("Изменено групп: {0}; исключено членств: {1} из {2}" -f $plan.Count, $totalToRemove, $totalCurrent) 'OK'
    }

    # ==========================================
    # ШАГ 3. Руководители подразделений (ManagedBy)
    # ==========================================
    Write-Log "Шаг 3: руководители подразделений" 'STEP'
    foreach ($m in $excelManagers) {
        $deptKey = Normalize-Name $m.Department
        if (-not $usersByEmpNum.ContainsKey($m.EmployeeNumber)) { continue }
        if (-not $managedGroupsMap.ContainsKey($deptKey)) {
            Write-Log "Группа подразделения '$($m.Department)' не найдена" 'WARN'
            continue
        }
        $manager = $usersByEmpNum[$m.EmployeeNumber]
        $group   = $managedGroupsMap[$deptKey]
        if ($PSCmdlet.ShouldProcess($group.Name, "Set ManagedBy=$($manager.SamAccountName)")) {
            try {
                Set-ADGroup -Identity $group -ManagedBy $manager.DistinguishedName -Server $DcServer
            } catch {
                Write-Log "Set ManagedBy fail $($group.Name): $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 4. Карантин и удаление групп
    # ==========================================
    # Группа, исчезнувшая из 1С, не удаляется сразу: сначала карантин на
    # $QuarantineDays дней со скрытием из адресной книги. Это даёт время распознать
    # переименование, а главное — сохраняет SID, на котором держатся доступы:
    # группы рассылки вложены в группы Jira/Confluence, и удаление рвёт эту связь.
    Write-Log "Шаг 4: карантин и удаление групп, отсутствующих в 1С" 'STEP'
    $finalGroups = Remove-IgnoredGroups (Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                               -SearchBase $TargetOU -Server $DcServer `
                               -Properties Member, MemberOf, mail, Name, SamAccountName, info, objectSid)

    foreach ($g in $finalGroups) {
        $key         = Normalize-Name $g.Name
        $inExcel     = $excelGroupKeys.ContainsKey($key)
        $quarantined = if ($g.info -and ($g.info -match 'QUARANTINE:(\d{4}-\d{2}-\d{2})')) { $matches[1] } else { $null }

        # Группа вернулась в справочник — карантин снимаем.
        if ($inExcel) {
            if ($quarantined) {
                if ($PSCmdlet.ShouldProcess($g.Name, "Снять карантин")) {
                    try {
                        Set-ADGroup -Identity $g -Clear info -Server $DcServer
                        Set-DistributionGroup -Identity $g.SamAccountName `
                            -HiddenFromAddressListsEnabled $false -DomainController $DcServer
                        Add-HistoryEvent -EventName 'Unquarantine' -Group $g.Name -Details @{ Sam = $g.SamAccountName }
                        Write-Log "Карантин снят: $($g.Name)" 'OK'
                    } catch { Write-Log "Снятие карантина не удалось $($g.Name): $($_.Exception.Message)" 'ERROR' }
                }
            }
            continue
        }

        # Защищённые в protected_groups.json не трогаем вовсе.
        $email = if ($g.mail) { $g.mail.ToLowerInvariant().Trim() } else { $null }
        if ($email -and $protectedByEmail.ContainsKey($email)) {
            Write-Log "Защищена JSON, пропуск: $($g.Name) ($email)" 'WARN'
            continue
        }

        if (-not $quarantined) {
            if ($PSCmdlet.ShouldProcess($g.Name, "Карантин")) {
                try {
                    Set-ADGroup -Identity $g -Replace @{ info = ("QUARANTINE:{0:yyyy-MM-dd}" -f (Get-Date)) } -Server $DcServer
                    Set-DistributionGroup -Identity $g.SamAccountName `
                        -HiddenFromAddressListsEnabled $true -DomainController $DcServer
                    Add-HistoryEvent -EventName 'Quarantine' -Group $g.Name -Details @{
                        Sam = $g.SamAccountName
                        Sid = if ($g.objectSid) { $g.objectSid.Value } else { '' }
                        Members = @($g.Member).Count
                        # Вхождения в чужие группы записываем до удаления: если доступ
                        # всё же будет потерян, по этой записи его можно вернуть.
                        MemberOf = @($g.MemberOf)
                    }
                    Write-Log "В карантин на $QuarantineDays дн.: $($g.Name)" 'WARN'
                } catch { Write-Log "Карантин не удался $($g.Name): $($_.Exception.Message)" 'ERROR' }
            }
            continue
        }

        $daysInQuarantine = ((Get-Date) - [datetime]$quarantined).TotalDays
        if ($daysInQuarantine -lt $QuarantineDays) {
            Write-Log ("В карантине {0:N0} из {1} дн.: {2}" -f $daysInQuarantine, $QuarantineDays, $g.Name) 'WARN'
            continue
        }

        # Карантин истёк, а кандидата на переименование так и не нашлось: подразделение
        # действительно расформировано. Размораживаем состав — держать доступ через
        # мёртвое подразделение нельзя, а как улика он больше не нужен.
        if ($g.Member -and $g.Member.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($g.Name, "Remove $(@($g.Member).Count) members")) {
                try {
                    Remove-ADGroupMember -Identity $g -Members $g.Member -Server $DcServer -Confirm:$false
                    Write-Log "Состав разморожен и очищен перед удалением: $($g.Name)" 'WARN'
                } catch {
                    Write-Log "Не удалось очистить состав $($g.Name): $($_.Exception.Message)" 'ERROR'
                    continue
                }
            }
        }

        if ($PSCmdlet.ShouldProcess($g.Name, "Disable + Remove")) {
            # Пишем всё, что понадобится для восстановления: SID, состав и вхождения
            # в чужие группы (Jira/Confluence). После удаления этих данных больше нет.
            Add-HistoryEvent -EventName 'Delete' -Group $g.Name -Details @{
                Sam = $g.SamAccountName
                Sid = if ($g.objectSid) { $g.objectSid.Value } else { '' }
                Mail = $g.mail; QuarantinedSince = $quarantined
                MemberOf = @($g.MemberOf); Members = @($g.Member)
            }
            try {
                Disable-DistributionGroup -Identity $g.DistinguishedName -Confirm:$false -DomainController $DcServer
            } catch { Write-Log "Disable fail $($g.Name): $($_.Exception.Message)" 'WARN' }
            try {
                Remove-ADGroup -Identity $g -Confirm:$false -Server $DcServer
                Write-Log "Удалена после карантина: $($g.Name)" 'OK'
            } catch { Write-Log "Remove fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
        }
    }

    # ==========================================
    # Реестр групп: собственный источник постоянства
    # ==========================================
    # 1С стабильного идентификатора не даёт, поэтому история названий ведётся здесь.
    # Ключ — SamAccountName: он не меняется никогда и переживает переименования.
    Write-Log "Обновление реестра групп" 'STEP'
    $stamp = (Get-Date).ToString('yyyy-MM-dd')
    $liveSams = @{}
    foreach ($g in (Remove-IgnoredGroups (Get-ADGroup -LDAPFilter "(description=$GroupDesc)" -SearchBase $TargetOU -Server $DcServer `
                                -Properties Name, SamAccountName, mail, info, objectSid))) {
        $liveSams[$g.SamAccountName] = $true
        $entry = $registry[$g.SamAccountName]
        if ($null -eq $entry) {
            $entry = [PSCustomObject]@{
                Sam = $g.SamAccountName; Sid = ''; Mail = ''; CurrentName = ''
                Names = @(); FirstSeen = $stamp; LastSeen = $stamp; Status = 'Active'
            }
            $registry[$g.SamAccountName] = $entry
        }

        # Смена названия фиксируется в истории: у прежнего проставляется дата конца.
        if ($entry.CurrentName -ne $g.Name) {
            $names = New-Object System.Collections.Generic.List[object]
            foreach ($n in @($entry.Names)) {
                if ($n.Name -and -not $n.To -and $n.Name -ne $g.Name) {
                    $n.To = $stamp
                }
                $names.Add($n)
            }
            if (-not ($names | Where-Object { $_.Name -eq $g.Name -and -not $_.To })) {
                $names.Add([PSCustomObject]@{ Name = $g.Name; From = $stamp; To = $null })
            }
            # .ToArray(), а не @(): см. комментарий в Write-JsonArray.
            $entry.Names       = $names.ToArray()
            $entry.CurrentName = $g.Name
        }

        $entry.Sid      = if ($g.objectSid) { $g.objectSid.Value } else { $entry.Sid }
        $entry.Mail     = $g.mail
        $entry.LastSeen = $stamp
        $entry.Status   = if ($g.info -match 'QUARANTINE:') { 'Quarantine' } else { 'Active' }
    }
    foreach ($sam in @($registry.Keys)) {
        if (-not $liveSams.ContainsKey($sam)) { $registry[$sam].Status = 'Deleted' }
    }
    Write-JsonArray -Path $RegistryJsonPath -Items $registry.Values
    Write-Log "Реестр: $($registry.Count) записей -> $RegistryJsonPath" 'OK'
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
    throw
}
finally {
    # Журнал пишем всегда, в том числе после аварии: события до сбоя нужны для разбора.
    try {
        if ($script:History.Count -gt 0) {
            Write-JsonArray -Path $HistoryJsonPath -Items $script:History
            Write-Log "Журнал событий: $($script:History.Count) записей -> $HistoryJsonPath" 'OK'
        }
    } catch { Write-Log "Не удалось записать журнал: $($_.Exception.Message)" 'ERROR' }

    if ($session) {
        Remove-PSSession $session -ErrorAction SilentlyContinue
        Write-Log "Сессия Exchange закрыта." 'OK'
    }
    Write-Log "Готово (DC=$DcServer)." 'STEP'
    Stop-Transcript | Out-Null
}
