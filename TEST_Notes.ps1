Import-Module PCXLab.Excel -Force
Import-Module PCXLab.Core -Force
Get-Command -Module PCXLab.Excel
Get-Command -Module PCXLab.Core

Remove-Module PCXLab.Excel -Force
Remove-Module PCXLab.Core -Force
Get-Command -Module PCXLab.Excel
Get-Command -Module PCXLab.Core

Get-Module PCXLab.Excel -ListAvailable
Get-Module PCXLab.Core -ListAvailable

Import-Module PCXLab.Excel -Force


cd C:\Projects\Automation\Modules

Import-Module .\PCXLab.Core
Get-Command -Module PCXLab.Core

Import-Module .\PCXLab.Excel
Get-Command -Module PCXLab.Excel

Get-Module PCXLab.Excel -ListAvailable
Get-Module PCXLab.Core -ListAvailable


$env:PSModulePath -split ";"

$env:PSModulePath += ";C:\Projects\Automation\Modules"
$env:PSModulePath -split ";"

$env:PSModulePath += ";C:\Projects\Automation\Modules"

Get-Module PCXLab.Excel -ListAvailable
Get-Module PCXLab.Core -ListAvailable

$env:PSModulePath += ";C:\Projects\Automation\Modules"

Get-Module PCXLab.Excel -ListAvailable
Get-Module PCXLab.Core -ListAvailable

<#

🧠 When do you need permanent setup?

Only if you want:

👉 Use modules anywhere (outside your script)

Then you can install modules into:

C:\Users\Administrator\Documents\WindowsPowerShell\Modules

#>