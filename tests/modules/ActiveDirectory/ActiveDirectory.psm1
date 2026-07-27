# Поддельный модуль ActiveDirectory: каталог в памяти.
# Нужен, чтобы прогнать настоящий Sync-DistributionGroups.ps1 без домена.
# Хранилище — $global:ADStore, его наполняет тест.

function script:Copy-Obj {
    param($o)
    if ($null -eq $o) { return $null }
    $h = [ordered]@{}
    foreach ($p in $o.PSObject.Properties) {
        if ($p.Value -is [System.Collections.IList]) { $h[$p.Name] = @($p.Value) }
        else { $h[$p.Name] = $p.Value }
    }
    return [PSCustomObject]$h
}

function script:Find-Group {
    param($Identity)
    $id = if ($Identity -is [string]) { $Identity } else { $Identity.DistinguishedName }
    if (-not $id -and $Identity.SamAccountName) { $id = $Identity.SamAccountName }
    foreach ($g in $global:ADStore.Groups) {
        if ($g.DistinguishedName -eq $id -or $g.SamAccountName -eq $id -or $g.Name -eq $id) { return $g }
    }
    return $null
}

function Get-ADGroup {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]$Identity, $Filter, $LDAPFilter, $SearchBase, $Server,
        $Properties, $ResultPageSize, [switch]$ErrorActionSilent
    )
    $result = @($global:ADStore.Groups)

    if ($LDAPFilter -and ($LDAPFilter -match '\(description=(.+)\)')) {
        $desc = $matches[1]
        $result = @($result | Where-Object { $_.Description -eq $desc })
    }
    if ($Filter -and ($Filter -match "SamAccountName -eq '([^']+)'")) {
        $sam = $matches[1]
        $result = @($result | Where-Object { $_.SamAccountName -eq $sam })
    }
    if ($Filter -and ($Filter -match "DisplayName -eq '([^']+)'")) {
        $dn = $matches[1]
        $result = @($result | Where-Object { $_.DisplayName -eq $dn })
    }
    if ($Identity) {
        $g = script:Find-Group $Identity
        if (-not $g) { throw "Cannot find an object with identity: '$Identity'" }
        $result = @($g)
    }
    if ($SearchBase) { $result = @($result | Where-Object { $_.DistinguishedName -like "*$SearchBase" }) }

    # Каждый Get- отдаёт снимок, как в настоящем AD: последующие изменения
    # в каталоге не должны задним числом менять уже полученный объект.
    return @($result | ForEach-Object { script:Copy-Obj $_ })
}

function Get-ADUser {
    [CmdletBinding()]
    param($Identity, $Filter, $LDAPFilter, $SearchBase, $Server, $Properties, $ResultPageSize)
    $result = @($global:ADStore.Users)

    if ($Filter) {
        # Разбираем ровно те два фильтра, которые строит синхронизация.
        if ($Filter -match '-not \(EmployeeNumber') {
            $result = @($result | Where-Object { -not $_.EmployeeNumber })
        } elseif ($Filter -match 'EmployeeNumber -like') {
            $result = @($result | Where-Object { $_.EmployeeNumber })
        }
        if ($Filter -match "Enabled -eq 'True'") { $result = @($result | Where-Object { $_.Enabled }) }
        if ($Filter -match 'Department -like') { $result = @($result | Where-Object { $_.Department }) }
        $result = @($result | Where-Object {
            $_.SamAccountName -notlike 'adm-*' -and $_.SamAccountName -notlike 'vpn-*' -and
            $_.SamAccountName -notlike '*test*' -and $_.Name -notlike '*ADM*' -and
            $_.Name -notlike '*VPN*' -and $_.Name -notlike '*KZ*'
        })
    }
    if ($SearchBase) { $result = @($result | Where-Object { $_.DistinguishedName -like "*$SearchBase" }) }
    if ($Identity) { $result = @($result | Where-Object { $_.SamAccountName -eq $Identity -or $_.DistinguishedName -eq $Identity }) }
    return @($result | ForEach-Object { script:Copy-Obj $_ })
}

function New-ADGroup {
    [CmdletBinding()]
    param($Name, $SamAccountName, $GroupCategory, $GroupScope, $Path, $Description, $DisplayName, $Server)
    if (script:Find-Group $Name) { throw "The object already exists: $Name" }
    $global:ADStore.NextSid++
    $global:ADStore.Groups += [PSCustomObject]@{
        DistinguishedName = "CN=$Name,$Path"
        Name              = $Name
        SamAccountName    = $SamAccountName
        DisplayName       = $DisplayName
        Description       = $Description
        Member            = @()
        MemberOf          = @()
        ManagedBy         = $null
        mail              = $null
        info              = $null
        whenCreated       = $global:ADStore.Now
        objectSid         = [PSCustomObject]@{ Value = "S-1-5-21-1-1-1-$($global:ADStore.NextSid)" }
        HiddenFromAddressListsEnabled = $false
    }
}

function Set-ADGroup {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]$Identity, $DisplayName, $ManagedBy, $Replace, $Clear, $Server, $Description
    )
    $g = script:Find-Group $Identity
    if (-not $g) { throw "Cannot find an object with identity: '$Identity'" }
    if ($DisplayName) { $g.DisplayName = $DisplayName }
    if ($ManagedBy)   { $g.ManagedBy   = $ManagedBy }
    if ($Description) { $g.Description = $Description }
    if ($Replace)     { foreach ($k in $Replace.Keys) { $g.$k = $Replace[$k] } }
    if ($Clear) {
        foreach ($k in @($Clear)) {
            if ($null -eq $g.$k) { throw "The specified directory service attribute or value does not exist ($k)" }
            $g.$k = $null
        }
    }
}

function Rename-ADObject {
    [CmdletBinding()]
    param($Identity, $NewName, $Server)
    $g = script:Find-Group $Identity
    if (-not $g) { throw "Cannot find an object with identity: '$Identity'" }
    if (script:Find-Group $NewName) { throw "An attempt was made to add an object to the directory with a name that is already in use: $NewName" }
    $oldDn = $g.DistinguishedName
    $parent = $oldDn -replace '^CN=[^,]+,', ''
    $g.Name = $NewName
    $g.DistinguishedName = "CN=$NewName,$parent"
    # Связанные атрибуты в AD переживают переименование: ссылки на объект
    # хранятся по его идентификатору, а не по DN. Воспроизводим это.
    foreach ($other in $global:ADStore.Groups) {
        $other.Member   = @($other.Member   | ForEach-Object { if ($_ -eq $oldDn) { $g.DistinguishedName } else { $_ } })
        $other.MemberOf = @($other.MemberOf | ForEach-Object { if ($_ -eq $oldDn) { $g.DistinguishedName } else { $_ } })
    }
}

function Add-ADGroupMember {
    [CmdletBinding()]
    param($Identity, $Members, $Server)
    $g = script:Find-Group $Identity
    if (-not $g) { throw "Cannot find an object with identity: '$Identity'" }
    foreach ($m in @($Members)) {
        $dn = if ($m -is [string]) { $m } else { $m.DistinguishedName }
        if ($g.Member -contains $dn) { throw "The specified account name is already a member of the group" }
        $g.Member += $dn
        $child = script:Find-Group $dn
        if ($child) { $child.MemberOf += $g.DistinguishedName }
    }
}

function Remove-ADGroupMember {
    [CmdletBinding()]
    param($Identity, $Members, $Server, [switch]$Confirm)
    $g = script:Find-Group $Identity
    if (-not $g) { throw "Cannot find an object with identity: '$Identity'" }
    foreach ($m in @($Members)) {
        $dn = if ($m -is [string]) { $m } else { $m.DistinguishedName }
        $g.Member = @($g.Member | Where-Object { $_ -ne $dn })
        $child = script:Find-Group $dn
        if ($child) { $child.MemberOf = @($child.MemberOf | Where-Object { $_ -ne $g.DistinguishedName }) }
    }
}

function Remove-ADGroup {
    [CmdletBinding()]
    param($Identity, $Server, [switch]$Confirm)
    $g = script:Find-Group $Identity
    if (-not $g) { throw "Cannot find an object with identity: '$Identity'" }
    foreach ($other in $global:ADStore.Groups) {
        $other.Member   = @($other.Member   | Where-Object { $_ -ne $g.DistinguishedName })
        $other.MemberOf = @($other.MemberOf | Where-Object { $_ -ne $g.DistinguishedName })
    }
    $global:ADStore.Groups = @($global:ADStore.Groups | Where-Object { $_.DistinguishedName -ne $g.DistinguishedName })
}

Export-ModuleMember -Function Get-ADGroup, Get-ADUser, New-ADGroup, Set-ADGroup,
                              Rename-ADObject, Add-ADGroupMember, Remove-ADGroupMember, Remove-ADGroup
