$env:PSModulePath += ";C:\Projects\Automation\Modules"

Import-Module PCXLab.Excel -Force
Import-Module PCXLab.Excel -Verbose

Get-Command -Module PCXLab.Excel

New-ModuleManifest `
-Path "C:\Projects\Automation\Modules\PCXLab.Excel\PCXLab.Excel.psd1" `
-RootModule "PCXLab.Excel.psm1" `
-ModuleVersion "1.0.0" `
-Description "PCXLab Excel Automation Module" `
-Author "PCXLab"

Import-Module PCXLab.Excel -Force

$file = Get-Item "C:\TEST\ICICI_CC_HP_2604_APR26_AZ_NEWFORMAT_CCStatement_Past15-04-2026_ConvertedFromXls.xlsx"

$result = Convert-ICICIFormat -File $file

$result | Select-Object -First 5 | Format-Table

$raw = Import-Excel $file.FullName -NoHeader
$raw[0..20]


cd "C:\Projects\Automation\ICICI-Excel-Transform\"

.\main.ps1 -Folder "C:\TEST"

.\main.ps1 -Folder "C:\TEST" -OutputFolder "C:\TEST\Output"


# Run once per module
New-ModuleManifest `
-Path "C:\Projects\Automation\Modules\PCXLab.Core\PCXLab.Core.psd1" `
-RootModule "PCXLab.Core.psm1" `
-ModuleVersion "1.0.0" `
-Description "PCXLab Core Utilities (Logging, Config, etc.)" `
-Author "PCXLab"


# Add module path (temporary)
$env:PSModulePath += ";C:\Projects\Automation\Modules"



Remove-Module PCXLab.Excel -ErrorAction SilentlyContinue
Import-Module PCXLab.Excel -Force
Import-Module .\PCXLab.Excel -Force


$file = Get-Item "C:\TEST\HDFC_CC_HP_*.xls"

$result = Convert-HDFCFormat -File $file

$result | Select -First 5


$raw = Import-Excel $file.FullName -NoHeader
$headerIndex = Get-HDFCHeader -RawData $raw

$raw[$headerIndex..($headerIndex+5)]



$workingFile = Convert-XlsToXlsx -File $file
$raw = Import-Excel $workingFile.FullName -NoHeader

$headerIndex = Get-HDFCHeader -RawData $raw
$raw[$headerIndex..($headerIndex+5)]


$result = Convert-HDFCFormat -File $file
$result | Select -First 5

Remove-Module PCXLab.Excel -ErrorAction SilentlyContinue
Import-Module .\PCXLab.Excel -Force


$file = Get-Item "C:\TEST\HDFC_CC_HP_*.xls"

$result = Convert-HDFCFormat -File $file

$result | Select -First 5 | Format-Table

$result | Select -First 5 | Format-List


#Reload moudle
Remove-Module PCXLab.Excel -ErrorAction SilentlyContinue
Import-Module .\PCXLab.Excel -Force

$file = Get-Item "C:\TEST\HDFC_CC_HP_*.xls"

$result = Convert-HDFCFormat -File $file

$result | Select -First 5 | Format-List



$file = Get-Item "C:\TEST\HDFC_CC_*.xls"

$result = Convert-HDFCFormat -File $file

$outFile = "C:\TEST\TEST_OUTPUT.xlsx"

$result | Export-Excel -Path $outFile -AutoSize -BoldTopRow


cd C:\Projects\Automation
Remove-Module PCXLab.Excel -ErrorAction SilentlyContinue
Remove-Module PCXLab.Core -ErrorAction SilentlyContinue

Import-Module .\Modules\PCXLab.Core -Force
Import-Module .\Modules\PCXLab.Excel -Force



.\PCXLab.Finance.Transform\main.ps1 -Folder "C:\TEST"

Remove-Module PCXLab.Excel -ErrorAction SilentlyContinue
Import-Module PCXLab.Excel -Force

Get-Module PCXLab.Excel
Get-Module PCXLab.Core

Get-Module PCXLab.Excel -ListAvailable

New-ModuleManifest `
-Path "C:\Projects\Automation\Modules\PCXLab.Finance\1.0.0\PCXLab.Finance.psd1" `
-RootModule "PCXLab.Finance.psm1" `
-ModuleVersion "1.0.0" `
-Author "PCXLab" `
-Description "PCXLab Finance Statement Automation Tool"

cd C:\Projects\Automation
$env:PSModulePath += ";C:\Projects\Automation\Modules"

Import-Module PCXLab.Finance -Force

Invoke-PCXLabFinance -Folder C:\TEST


Remove-Module PCXLab.Finance -ErrorAction SilentlyContinue
Import-Module PCXLab.Finance -Force

Invoke-PCXLabFinance -Folder C:\TEST

tree
Get-ChildItem -Recurse

Remove-Module PCXLab.Finance -Force
Remove-Module PCXLab.Core -Force
Remove-Module PCXLab.Excel -Force

Import-Module .\Modules\PCXLab.Core -Force
Import-Module .\Modules\PCXLab.Excel -Force
Import-Module .\Modules\PCXLab.Finance -Force

Invoke-PCXLabFinance -Folder C:\TEST



tree /f


Remove-Module PCXLab.Finance -Force -ErrorAction SilentlyContinue
Remove-Module PCXLab.Excel -Force -ErrorAction SilentlyContinue
Remove-Module PCXLab.Core -Force -ErrorAction SilentlyContinue

Import-Module .\Modules\PCXLab.Core -Force
Import-Module .\Modules\PCXLab.Excel -Force
Import-Module .\Modules\PCXLab.Finance -Force

Invoke-PCXLabFinance -Folder C:\TEST

(Get-Module PCXLab.Core).Version
(Get-Command Write-Log).Source

Get-Module PCXLab* | Remove-Module -Force

$env:PSModulePath = "C:\Projects\Automation\Modules;" + $env:PSModulePath

cd C:\Projects\Automation

Import-Module PCXLab.Core -Force
Import-Module PCXLab.Excel -Force
Import-Module PCXLab.Finance -Force

Start-LogSession -LogFolder C:\TEST\logs
Write-Log "manual test"

Invoke-PCXLabFinance -Folder C:\TEST

Import-Module "C:\Projects\Automation\Modules\PCXLab.Core\dev\PCXLab.Core.psd1" -Force

# tag to be added for next release
git tag -a v1.2.2 -m "Release including CHANGELOG"
git push origin v1.2.2
