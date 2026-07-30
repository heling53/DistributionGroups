#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Однократное создание кастомной группы рассылки «Все сотрудники» (employees@transitcard.ru).
.DESCRIPTION
    «Все сотрудники» — верхнеуровневая группа рассылки, которой НЕТ в выгрузке 1С
    (Подразделения.xlsx). Она создаётся вручную ОДИН РАЗ этим скриптом, после чего на неё
    опирается Sync-DistributionGroups.ps1:
      - Шаг 3 вкладывает в неё отделы верхнего уровня (бывшие прямые подчинённые
        «Руководство») и саму группу «Руководство»;
      - Шаг 4 её не удаляет (она в списке «обязательно существующих»).

    Группе намеренно присваивается то же Description, что и у остальных управляемых групп
    ('Группа рассылки подразделения_v1.0'), чтобы синхронизация и аудит подхватывали её
    штатной логикой (очистка состава и перестроение иерархии при каждом прогоне).

    Скрипт идемпотентен: если группа с таким именем или адрес employees@ уже существуют,
    он ничего не пересоздаёт и не меняет, а только сообщает об этом.
.EXAMPLE
    .\New-AllEmployeesGroup.ps1 -WhatIf
    Показать, что будет сделано, без внесения изменений.
.EXAMPLE
    .\New-AllEmployeesGroup.ps1
    Создать группу «Все сотрудники» с адресом employees@transitcard.ru.
.NOTES
    Запускать под учётной записью с правами на создание групп в целевом OU и на управление
    Exchange (Enable-DistributionGroup). Все операции — на одном DC ($DcServer).
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$GroupName      = 'Все сотрудники',
    [string]$Alias          = 'employees',
    [string]$MailDomain     = 'transitcard.ru',
    [string]$PrimarySmtp    = 'employees@transitcard.ru',
    [string]$TargetOU       = 'OU=Группы рассылки,DC=transitcard,DC=ru',
    [string]$GroupDesc      = 'Группа рассылки подразделения_v1.0',
    [string]$ExchangeServer = 'srv-ex16-01.transitcard.ru',
    [string]$DcServer       = 'srv-dc3.transitcard.ru',
    # Для общекорпоративного списка можно запретить приём писем от внешних/неаутентифицированных
    # отправителей, выставив $true. По умолчанию $false — как у остальных групп подразделений.
    [bool]$RequireSenderAuthenticationEnabled = $false
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
} catch { }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) {
        'ERROR' { 'Red' }; 'WARN' { 'Yellow' }; 'OK' { 'Green' }; 'STEP' { 'Cyan' }; default { 'Gray' }
    }
    Write-Host ("[{0:HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message) -ForegroundColor $color
}

Import-Module ActiveDirectory -ErrorAction Stop

$session = $null
try {
    Write-Log "Подключение к Exchange $ExchangeServer..." 'STEP'
    $session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeServer) `
                             -Authentication Kerberos
    Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
    Write-Log "Сессия Exchange открыта." 'OK'

    # 1) Группа с таким именем уже существует в целевом OU?
    $existingByName = Get-ADGroup -Filter "Name -eq '$GroupName'" `
                                  -SearchBase $TargetOU -Server $DcServer -ErrorAction SilentlyContinue
    if ($existingByName) {
        Write-Log "Группа '$GroupName' уже существует ($($existingByName.DistinguishedName)). Создание не требуется." 'WARN'
        return
    }

    # 2) Адрес employees@ уже занят другой группой/объектом?
    $existingByMail = Get-ADGroup -LDAPFilter "(proxyAddresses=*$PrimarySmtp*)" `
                                  -Server $DcServer -ErrorAction SilentlyContinue
    if ($existingByMail) {
        Write-Log "Адрес $PrimarySmtp уже назначен группе '$($existingByMail.Name)'. Прерываю, чтобы не создать дубликат." 'ERROR'
        return
    }

    # 3) Подбор свободного SamAccountName (на случай, если '$Alias' уже занят).
    $sam = $Alias
    $i = 1
    while (Get-ADGroup -Filter "SamAccountName -eq '$sam'" -Server $DcServer -ErrorAction SilentlyContinue) {
        $sam = ($Alias + $i); $i++
    }

    if ($PSCmdlet.ShouldProcess($GroupName, "New-ADGroup + Enable-DistributionGroup ($PrimarySmtp)")) {
        New-ADGroup -Name $GroupName -SamAccountName $sam `
                    -GroupCategory Distribution -GroupScope Universal `
                    -Path $TargetOU -Description $GroupDesc -DisplayName $GroupName `
                    -Server $DcServer
        Write-Log "Создана AD-группа: $GroupName ($sam)" 'OK'

        Enable-DistributionGroup -Identity $sam -Alias $Alias `
            -PrimarySmtpAddress $PrimarySmtp -DomainController $DcServer
        Write-Log "Включена как группа рассылки: $PrimarySmtp" 'OK'

        Set-DistributionGroup -Identity $sam `
            -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled `
            -DomainController $DcServer
        Write-Log "RequireSenderAuthenticationEnabled = $RequireSenderAuthenticationEnabled" 'OK'

        Write-Log "Готово. Теперь Sync-DistributionGroups.ps1 будет наполнять '$GroupName' и вкладывать в неё «Руководство»." 'OK'
    }
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" 'ERROR'
    throw
}
finally {
    if ($session) {
        Remove-PSSession $session -ErrorAction SilentlyContinue
        Write-Log "Сессия Exchange закрыта." 'OK'
    }
}
