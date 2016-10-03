function Get-MPFromBundle {
    param(
        [System.IO.FileSystemInfo]$BundleFile
    )

    # Load the System Center SDK Assemblies
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.Core") | Write-Debug
    [Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.Packaging") | Write-Debug

    Write-Debug "Creating Management pack objects, mpstore, xmlwriter and mpbreader."
    $mpstore = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackFileStore($outputDir)
    $mpbReader = [Microsoft.EnterpriseManagement.Packaging.ManagementPackBundleFactory]::CreateBundleReader()

    $mpb = $mpbReader.Read($BundleFile.FullName,$mpstore)
    foreach ($mp in $mpb.ManagementPacks)
    {
        Write-Debug "MP found in bundle, $($mp.Name)."
        Write-Debug "mps array count: $($managementPacks.count)"
        $managementPacks += $mp
        Write-Debug "mps array count after add: $($managementPacks.count)"        
    }

    return $managementPacks

}