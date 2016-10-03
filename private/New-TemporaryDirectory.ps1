


function New-TemporaryDirectory {
    param(
        [string]$prefix
    )
    $parent = [System.IO.Path]::GetTempPath()    
    [string] $name = [System.Guid]::NewGuid()
    $name = "$($prefix)_$($name)"
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}   