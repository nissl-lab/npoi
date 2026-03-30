param([string]$NupkgPath)

$tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.Guid]::NewGuid().ToString())
[System.IO.Directory]::CreateDirectory($tempDir) | Out-Null

try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($NupkgPath, $tempDir)
    
    Get-ChildItem -Path $tempDir -Recurse -File | Where-Object { $_.Name -match '^NPOI\.Pack\.' } | Remove-Item -Force
    
    if (Test-Path $NupkgPath) {
        Remove-Item $NupkgPath -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $NupkgPath)
}
finally {
    [System.IO.Directory]::Delete($tempDir, $true)
}
