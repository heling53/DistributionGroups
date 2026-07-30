@{
    RootModule        = 'ActiveDirectory.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '9b1a2c3d-4e5f-6a7b-8c9d-0e1f2a3b4c5d'
    Author            = 'test'
    Description       = 'Поддельный ActiveDirectory для тестов'
    PowerShellVersion = '5.1'
    FunctionsToExport = @('Get-ADGroup','Get-ADUser','New-ADGroup','Set-ADGroup','Rename-ADObject','Add-ADGroupMember','Remove-ADGroupMember','Remove-ADGroup')
}
