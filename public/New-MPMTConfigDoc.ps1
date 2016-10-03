
function New-MPMTConfigDoc 
{
    <# 
    .SYNOPSIS 
    This command creates a config document for the management pack. It's an excel file with all the classes, monitors, 
    discoveries and rules defined in the management pack. 
    .DESCRIPTION 
    This command creates a config document for the management pack. It's an excel file with all the classes, monitors, 
    discoveries and rules defined in the management pack. 
    .EXAMPLE 
    Get-Item C:\temp\Microsoft.SQLServer.2014.Monitoring.mp | New-MPMTConfigDoc 
    .EXAMPLE 
    Get-Item C:\temp\Microsoft.SQLServer.2014.Monitoring.mp | New-MPMTConfigDoc -outputDir C:\temp\overrideMPs\ 
    .EXAMPLE 
    Get-ChildItem C:\temp\ -Filter *.mp | New-MPMTConfigDoc -outputDir C:\temp\overrideMPs\
    .PARAMETER ManagementPack
    The management pack to create a config document for. It can be a single file or a collection (via the pipeline).
    .PARAMETER outputDir
    The output folder where the config document is to be saved. If no outputDir is supplied it will be written to the current dir.
    #> 

    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]         
        [System.IO.FileSystemInfo]$ManagementPack,
        [string]$outputDir
    )

    begin {
        # If no output dir is supplied, set it to where the script is called from.
        if([string]::IsNullOrEmpty($outputDir)) {
            $outputDir = (Get-Location).Path
        }
        Write-Output "Output folder: $outputDir"

        # Creating a temporary directory to store files. Used when extacting mp files from an msi.
        $tempDir = New-TemporaryDirectory -Prefix "MMPT"
        Write-Debug "Creating temporary directory $($tempDir)."

        Write-Debug "Loading System Center SDK Assemblies:"
        # Load the System Center SDK Assemblies
        [Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.Core") | Write-Debug
        [Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.Packaging") | Write-Debug
    } # begin

    process {
        Write-Debug "Getting file object for mp '$ManagementPack'."
        Write-Debug "Gettype '$($ManagementPack.GetType())'."
        Write-Debug "FullName '$($ManagementPack.FullName)'."
        $file = $ManagementPack     
        Write-Debug "Supplied file: '$($file.Name)'."

        $managementPacks =  @() # This is an array that will contain the management packs. In case of a mp 
                                # bundle there can be more than one mp supplied. We need a override mp for
                                # each of those. But in most cases this will only contain one mp. 

        # Check the file extension
        Write-Debug "File extension: '$($file.Extension)'"
        switch ($file.Extension) {
            ".msi" {
                Write-Debug "MSI file found."        
                Write-Debug "Unpacking to temp directory: '$($tempDir)'."
                Write-Debug "Command: & msiexec ""/a"" $($file.FullName) ""/qb"" ""TARGETDIR=$tempDir"""
                & msiexec "/a" $($file.FullName) "/qb" "TARGETDIR=$tempDir" | Out-Null

                 # Get all management packs
                $managementPackFiles = Get-ChildItem -Path "$($tempDir)" -Include @("*.mp") -Recurse
                Write-Debug "Unpacked management pack files: $($managementPackFiles.count)"

                # Creating MP object for each file and add to collection
                Write-Debug "Adding extracted MPs to managementPacks collection"
                foreach ($mpFile in $managementPackFiles) {
                    Write-Debug "`t$($mpFile.Name)"
                    $managementPacks += New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPack($mpFile.FullName)
                }
                
                # Get all Bundles and add the MPs to the collection
                $bundlesFiles = Get-ChildItem -Path "$($tempDir)" -Include @("*.mpb") -Recurse
                
                foreach ($bundle in $bundlesFiles) {
                    Write-Debug "MP Bundle file found, $($bundle.Name)."
                    $MPsInBundle = Get-MPFromBundle $bundle
                    $managementPacks += $MPsInBundle
                }

                # Check that we found any management packs.
                if ($managementPacks.count -lt 1) 
                {
                    Write-Warning "No management packs found in '$tempDir'. Verify the folder name and that the files have been copied to the correct folder."
                    break
                }
            }
            ".mp" {
                Write-Debug "MP file found."
                $managementPacks += New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPack($file.FullName)
            }
             ".mpb" {
                Write-Debug "MP Bundle file found."
                $MPsInBundle = Get-MPFromBundle -BundleFile $file
                $managementPacks += $MPsInBundle
            }
            ".xml" {
                Write-Warning "Unsealed management packs are not supported."
                break
            }
            Default {
                Write-Warning "This file type is not supported. It must be either a mp file or a msi file."
                break
            }
        }

        # Check that we found any files
        Write-Debug "Found $($managementPacks.Count) management pack(s)."

        # Create management packs
        foreach ($mp in $managementPacks) {
            Write-Output "Creating config document for '$($mp.Name)'."
            New-ConfigurationDocument -file $mp -targetFolder $outputDir
        }
    } # process

    end {
        # Deleting temporary directory
        if ($tempDir -ne $null) {
            if (Test-Path $tempDir) {
                Write-Debug "Removing temporary directory and files, '$($tempDir)'."
                Write-Debug "Command: 'Remove-Item $tempDir -Recurse -Force'"
                try {
                    Remove-Item $tempDir -Recurse -Force -ErrorAction Stop
                }
                catch [System.IO.IOException] {
                    # TODO: Solve this somehow!!!
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    Write-Warning "Error deleting temporary folder: '$tempDir'.`nError message: '$ErrorMessage'`n`nFile and folder needs to be manually deleted."
                }
            }
        }
        Write-Output "Done"
    } # end

} # function