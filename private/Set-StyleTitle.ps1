function Set-StyleTitle
{
    param($cell)

    Write-Debug "Set-StyleTitle: cell : '$($cell.Value2)'"

    $cell.Font.Size = 18
    $cell.Font.Name = "Calibri Light"

}