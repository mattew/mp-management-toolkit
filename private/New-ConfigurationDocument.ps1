function New-ConfigurationDocument
{
    # This function reads and documents a management pack 
    # in an excel file.
    param($file, $targetFolder)

    Write-Debug "New-ConfigurationDocument: Supplied file: '$($file.Name)'."
    Write-Debug "New-ConfigurationDocument: Extension: '$($file.Extension)'."
    Write-Debug "New-ConfigurationDocument: Configuration document will be written to '$targetFolder'."    
    
    # Variables
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.Core") | Out-Null
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.Packaging") | Out-Null
    $mps = @()
    $mpstore = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackFileStore($mpStorePath)
    $mpbReader = [Microsoft.EnterpriseManagement.Packaging.ManagementPackBundleFactory]::CreateBundleReader()
    # Excel objects/variables
    # Constants for alignment
    $EXCELRIGHTALIGNMENT = -4152
    $EXCELLEFTALIGNMENT = -4131
    $EXCELCENTERTALIGNMENT = -4108
    # Something hackety for excel to work
    $newci = [System.Globalization.CultureInfo]"en-US"
    [system.threading.Thread]::CurrentThread.CurrentCulture = $newci
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $newci
    # Workbook template variables
    $summarySheet = 1
    $classesSheet = 2
    $discoveriesSheet = 3
    $monitorsSheet = 4
    $rulesSheet = 5
    $tasksSheet = 6
    $SheetsInNewWorkbook = 6
   
    # The Excel application object and workbook.
    $excel = New-Object -ComObject Excel.Application # This creates a VERBOSE message which I can't turn of.
    $excel.SheetsInNewWorkbook = $SheetsInNewWorkbook
    $excel.DisplayAlerts = $false
    $excel.Visible = $false
    #$excel.Visible = $true

    # Create a new workbook
    $workbook = $excel.Workbooks.Add()



    # Create the summary sheet.
    Write-Debug "Create the Summary sheet."
    $summarySheet = $workbook.Worksheets.Item($summarySheet)
    $summarySheet.Name = "MP Summary"
    $summarySheet.Activate()

    # Create the headers column
    Write-Debug "Create the headers column."
    $headers = @( "MP Name"
                , "MP ID"
                , "MP File/Bundle"
                , "Version"
                , "Nr of classes"
                , "Nr of discoveries"
                , "Nr of monitors"
                , "Nr of rules"
                , "Nr of tasks"
                )            
    Write-Headers $workbook $headers 'Vertical'


    # Create the values column
    Write-Debug "Create the values column."
    $startRow = 3
    $currentRoW = $startRow
    $startColumn = 2
    $values = @( $mp.DisplayName
                , $mp.Name
                , $file.Name
                , $mp.Version
                , $($mp.GetClasses()).Count
                , $($mp.GetDiscoveries()).Count
                , $($mp.GetMonitors()).Count
                , $($mp.GetRules()).Count
                , $($mp.GetTasks()).Count
                )
    foreach ($value in $values)
    {
        $cell = $summarySheet.Cells.Item($currentRoW,$startColumn)
        $cell.Value2 = $value
        $currentRoW++
    }
    $summarySheet.Range("B$startRow","B$($currentRoW)").HorizontalAlignment = $EXCELRIGHTALIGNMENT
    # Autofit all column/rows
    $objRange = $summarySheet.UsedRange
    [void] $objRange.EntireColumn.Autofit()
    # This is after the autofit, else the first column becomes super wide.
    $summarySheet.Cells.Item(1,1) = $mp.DisplayName
    Set-StyleTitle $summarySheet.Cells.Item(1,1)



    # Create the Classes sheet
    Write-Debug "Create the Classes sheet."
    $classesSheet= $workbook.Worksheets.Item($classesSheet)
    $classesSheet.Activate()
    $classesSheet.Name = "Classes"
    $headers = @( "Name"
                , "Display Name"
                , "Base Class"
                , "Description"
                )
    Write-Headers $workbook $headers "Horizontal"


    $entities = $mp.GetClasses()
    $values = New-Object 'object[,]' $entities.Count,4
    for ($i = 0; $i -lt $entities.Count;$i++)
    {
        $values[$i,0] = $entities[$i].Name
        $values[$i,1] = $entities[$i].DisplayName
        $values[$i,2] = $($entities[$i].Base).Name
        $values[$i,3] = $entities[$i].Description
    }
    Write-ValuesToRange $workbook $values
    # Autofit all column/rows
    $objRange = $classesSheet.UsedRange
    [void] $objRange.EntireColumn.Autofit()  
    # This is after the autofit, else the first column becomes super wide.
    $classesSheet.Cells.Item(1,1) = "Classes"        
    Set-StyleTitle $classesSheet.Cells.Item(1,1)




    # Create the Discoveries sheet
    Write-Debug "Create the Discoveries sheet."
    $discoveriesSheet= $workbook.Worksheets.Item($discoveriesSheet)
    $discoveriesSheet.Activate()
    $discoveriesSheet.Name = "Discoveries"
    $headers = @( "Name"
                , "Display Name"
                , "Enabled"
                , "Target"
                , "Description"    
                )
    Write-Headers $workbook $headers "Horizontal"


    $entities = $mp.GetDiscoveries()
    $values = New-Object 'object[,]' $entities.Count,5
    for ($i = 0; $i -lt $entities.Count;$i++)
    {
        $values[$i,0] = $entities[$i].Name        
        $values[$i,1] = $entities[$i].DisplayName                        
        $values[$i,2] = $($($entities[$i].Enabled).ToString()).ToLower()
        $values[$i,3] = $($($entities[$i]).Target).Name
        $values[$i,4] = $entities[$i].Description
    }
    Write-ValuesToRange $workbook $values
    # Autofit all column/rows
    $objRange = $discoveriesSheet.UsedRange
    [void] $objRange.EntireColumn.Autofit()  
    # This is after the autofit, else the first column becomes super wide.
    $discoveriesSheet.Cells.Item(1,1) = "Discoveries"
    Set-StyleTitle $discoveriesSheet.Cells.Item(1,1)




    # Create the Monitors sheet
    Write-Debug "Create the Monitors sheet."
    $monitorsSheet= $workbook.Worksheets.Item($monitorsSheet)
    $monitorsSheet.Activate()
    $monitorsSheet.Name = "Monitors"
    $headers = @( "Name"
                , "Display Name"
                , "Enabled"
                , "Target"
                , "Description"    
                )
    Write-Headers $workbook $headers "Horizontal"
    $entities = $mp.GetMonitors()
    $values = New-Object 'object[,]' $entities.Count,5
    for ($i = 0; $i -lt $entities.Count;$i++)
    {
            if ($entities[$i].XmlTag -eq "UnitMonitor") 
            { 
                $values[$i,0] = $entities[$i].Name        
                $values[$i,1] = $entities[$i].DisplayName                        
                $values[$i,2] = $($($entities[$i].Enabled).ToString()).ToLower()
                $values[$i,3] = $($($entities[$i]).Target).Name
                $values[$i,4] = $entities[$i].Description
        }
    }
    Write-ValuesToRange $workbook $values
    # Autofit all column/rows
    $objRange = $monitorsSheet.UsedRange
    [void] $objRange.EntireColumn.Autofit()  
    # This is after the autofit, else the first column becomes super wide.
    $monitorsSheet.Cells.Item(1,1) = "Monitors"
    #$monitorsSheet.Cells.Item(1,1).Style = "Title"   
    Set-StyleTitle $monitorsSheet.Cells.Item(1,1)




    # Create the Rules sheet
    Write-Debug("Create the Rules sheet")
    $rulesSheet= $workbook.Worksheets.Item($rulesSheet)
    $rulesSheet.Activate()
    $rulesSheet.Name = "Rules"
    $headers = @( "Name"
                , "Display Name"
                , "Enabled"
                , "Target"
                , "Description"    
                )
    Write-Headers $workbook $headers "Horizontal"
    $entities = $mp.GetRules()
    $values = New-Object 'object[,]' $entities.Count,5
    for ($i = 0; $i -lt $entities.Count;$i++)
    {
            $values[$i,0] = $entities[$i].Name        
            $values[$i,1] = $entities[$i].DisplayName                        
            $values[$i,2] = $($($entities[$i].Enabled).ToString()).ToLower()
            $values[$i,3] = $($($entities[$i]).Target).Name
            $values[$i,4] = $entities[$i].Description
    }
    Write-ValuesToRange $workbook $values
    # Autofit all column/rows
    $objRange = $rulesSheet.UsedRange
    [void] $objRange.EntireColumn.Autofit()  
    # This is after the autofit, else the first column becomes super wide.
    $rulesSheet.Cells.Item(1,1) = "Rules"
    Set-StyleTitle $rulesSheet.Cells.Item(1,1)




    # Create the Tasks sheet
    Write-Debug "Create the Tasks sheet."
    $tasksSheet= $workbook.Worksheets.Item($tasksSheet)
    $tasksSheet.Activate()
    $tasksSheet.Name = "Tasks"
    $headers = @( "Name"
                , "Display Name"
                , "Enabled"
                , "Target"
                , "Description"    
                )
    Write-Headers $workbook $headers "Horizontal"


    $entities = $mp.GetTasks()
    $values = New-Object 'object[,]' $entities.Count,5
    for ($i = 0; $i -lt $entities.Count;$i++)
    {
        $values[$i,0] = $entities[$i].Name
        $values[$i,1] = $entities[$i].DisplayName
        $values[$i,2] = $($($entities[$i].Enabled).ToString()).ToLower()
        $values[$i,3] = $($($entities[$i]).Target).Name
        $values[$i,4] = $entities[$i].Description
    }
    Write-ValuesToRange $workbook $values
    # Autofit all column/rows
    $objRange = $tasksSheet.UsedRange
    [void] $objRange.EntireColumn.Autofit()  
    # This is after the autofit, else the first column becomes super wide.
    $tasksSheet.Cells.Item(1,1) = "Tasks"        
    Set-StyleTitle $tasksSheet.Cells.Item(1,1)


    # Activate the summary sheet
    $summarySheet.Activate()


    # Save file
    $xlsFileName = "$($targetFolder)\$($mp.Name).xls"
    $workbook.SaveAs($xlsFileName, 1)    
    $workbook.close()    

    Remove-Variable -Name excel
    Remove-Variable -Name workbook    
    Remove-Variable -Name mpstore
    Remove-Variable -Name mpbReader


    # Killing all Excel process
    Write-Debug "New-ConfigurationDocument: Killing all Excel process"
    Get-Process -Name Excel | Stop-Process
}
