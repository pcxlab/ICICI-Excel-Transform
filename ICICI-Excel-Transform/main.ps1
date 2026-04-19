param(
    [Parameter(Mandatory)]
    [string]$Folder,

    [string]$OutputFolder
)

# 🔹 Setup module path FIRST
$modulePath = Join-Path $PSScriptRoot "Modules"
$env:PSModulePath = "$modulePath;$env:PSModulePath"

# Add module path (temporary)
$env:PSModulePath += ";C:\Projects\Automation\Modules"

# 🔹 Import modules BEFORE logging
Import-Module PCXLab.Core -Force
Import-Module PCXLab.Excel -Force

# 🔹 Start logging
Start-LogSession -LogFolder (Join-Path $Folder "logs")

# 🔹 Ensure NuGet provider
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {

        #Write-Host "NuGet not found. Installing..." -ForegroundColor Yellow
        Write-Log "NuGet not found. Installing..."

        Install-PackageProvider -Name NuGet `
            -MinimumVersion 2.8.5.201 `
            -Force `
            -Scope CurrentUser `
            -ErrorAction Stop

        #Write-Host "NuGet installed successfully." -ForegroundColor Green
        Write-Log "NuGet installed successfully." "SUCCESS"
    }
}
catch {
    #Write-Host "Failed to install NuGet: $($_.Exception.Message)" -ForegroundColor Red
    Write-Log "Failed to install NuGet: $($_.Exception.Message)" "ERROR"
    exit
}

# 🔹 Ensure ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {

    #Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Write-Log "ImportExcel module not found. Installing..."

    try {
        Install-Module ImportExcel `
            -Scope CurrentUser `
            -Force `
            -AllowClobber `
            -ErrorAction Stop

        #Write-Host "ImportExcel installed successfully." -ForegroundColor Green
        Write-Log "ImportExcel installed successfully." "SUCCESS"
    }
    catch {
        #Write-Host "Failed to install ImportExcel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Failed to install ImportExcel: $($_.Exception.Message)" "ERROR"
        exit
    }
}

# 🔹 Validate input folder
if (-not (Test-Path $Folder)) {
    Write-Log "Input folder does not exist: $Folder" "ERROR"
    exit
}

# 🔹 Default output folder
if (-not $OutputFolder) {
    $OutputFolder = $Folder
}

# 🔹 Ensure output folder exists
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# 🔹 Process files
$files = Get-ChildItem $Folder -File

foreach ($file in $files) {

    if ($file.Name -match "_ConvertedFromXls" -or $file.Name -match "_Transformed") {
        Write-Log "Skipping: $($file.Name)"
        continue
    }

    Write-Log "Processing: $($file.Name)"

    try {
        $workingFile = Convert-XlsToXlsx -File $file
        $result = Convert-ICICIFormat -File $workingFile

        $outFileName = Get-OutputFileName `
            -File $file `
            -Converted:$($file.Extension -eq ".xls") `
            -Transformed

        $outFile = Join-Path $OutputFolder $outFileName

        $result | Export-Excel -Path $outFile -AutoSize -BoldTopRow

        Write-Log "Saved: $outFile" "SUCCESS"
    }
    catch {
        #Write-Host "Error processing $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error processing $($file.Name): $($_.Exception.Message)" "ERROR"
    }
}