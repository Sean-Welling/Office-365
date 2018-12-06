#############################################################
# AUTHOR  : Sean Welling
# REVISION: 1.0
# DATE    : 11/16/2018
# DESCR   : Assign Office 365 Licenses in Bulk
#############################################################

<#___________________________________________INCLUDE FUNCTIONS_____________________________________________#>

."\\karina\Information Technology\Scripts\Office 365\Office365-FUNCTIONS.ps1"

<#___________________________________________TRIM CSV HEADERS______________________________________________#>

$SourcePath = "\\karina\Information Technology\Scripts\Office 365\Office365Users.csv"

$cred = Get-Credential swelling@moldedfiberglass.com
Connect-AzureAD -Credential $cred

# Source Headers Dirty (source CSV has unwanted spaces)
$SourceHeadersDirty = Get-Content -Path $SourcePath -First 2 | ConvertFrom-Csv
 
# Source Headers Cleaned (removed spaces and :)
$SourceHeadersCleaned = $SourceHeadersDirty.PSObject.Properties.Name.Trim(' ') -Replace '\W', ''
 
# Import-CSV
$CSV = Import-CSV -Path $SourcePath -Header $SourceHeadersCleaned | Select-Object -Skip 1

$Logfile = "\\karina\Information Technology\Scripts\Office 365\Office365UserErrorLog.log"

<#______________________________________________CODE START_________________________________________________#>

$CSV | ForEach-Object {
    $Username = $_.Username.trim()
    $Office365 = $_.Office365.trim()
    $Project = $_.Project.trim()
    $Visio = $_.Visio.trim()
    $UserPrincipalName = "$Username@moldedfiberglass.com"

    if ($Office365 -eq "Y") {
        Add-O365License($UserPrincipalName)
    }

    if ($Project -eq "Y") {
        Add-ProjectLicense($UserPrincipalName)
    }

    if ($Visio -eq "Y") {
        Add-VisioLicense($UserPrincipalName)
    }


}