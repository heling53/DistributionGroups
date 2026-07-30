# Поддельный ImportExcel: возвращает строки, подготовленные тестом в $global:ExcelRows.
function Import-Excel {
    [CmdletBinding()]
    param($Path, $WorksheetName, $StartRow)
    return @($global:ExcelRows)
}
Export-ModuleMember -Function Import-Excel
