function Write-Headers
{
    param($workbook, $headers, $order = "Horizontal")

    Write-Debug "Write-Headers: workbook : '$($workbook)'"
    Write-Debug "Write-Headers: order: '$($order)'."
    Write-Debug "Write-Headers: headers: '$($headers)'."
    Write-Debug "Write-Headers: headers count: '$($headers.count)'."

    $startRow = 3
    $startColumn = 1

    if ($order -eq 'Horizontal')
    {
        $startcell = "$(Convert-NumberToA1 $startColumn)$startRow"
        $endcell = "$(Convert-NumberToA1 $($startColumn + $($headers.count-1)))$($startRow)"
        $range = $workbook.ActiveSheet.Range("$startcell","$endcell")
        $array = New-Object 'object[,]' 1,$headers.count
        for ($i = 0; $i -lt $headers.count; $i++)
        {
            $array[0,$i] = $headers[$i]
        }        
    }
    if ($order -eq 'Vertical')
    {
        $startcell = "$(Convert-NumberToA1 $startColumn)$startRow"
        $endcell = "$(Convert-NumberToA1 $startColumn)$($startRow + $($headers.count-1))"
        $range = $workbook.ActiveSheet.Range("$startcell","$endcell")

        $array = New-Object 'object[,]' $headers.count,1
        for ($i = 0; $i -lt $headers.count; $i++)
        {
            $array[$i,0] = $headers[$i]
        }        
    }

    Write-Debug "Write-Headers: Range = $($startcell):$($endcell)"
    
    $range.Value2 = $array
    $range.Font.Bold = $True 
    $range.Font.Name = "Calibri Light"
    $range.Font.Size = 11
}