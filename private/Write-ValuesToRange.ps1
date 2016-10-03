function Write-ValuesToRange
{
    param($workbook, $values)

    Write-Debug "Write-ValuesToRange: workbook : '$($workbook)'"
    Write-Debug "Write-ValuesToRange: values count: '$($values.count)'."

    $startRow = 4
    $startColumn = 1

    $startcell = "$(Convert-NumberToA1 $startColumn)$startRow"
    $endcell = "$(Convert-NumberToA1 $($startColumn + $values.GetLength(1) - 1))$($startRow + $values.GetLength(0) - 1)"
    Write-Debug "Write-ValuesToRange: Range = $($startcell):$($endcell)" 

    $range = $workbook.ActiveSheet.Range("$startcell","$endcell")
    $range.Value2 = $values
}