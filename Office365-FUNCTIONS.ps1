function Add-O365License($UserPrincipalName) {

    $ENTERPRISEPACK = '6fd2c87f-b296-42f0-b197-1e91e994b900'
    $EMS = 'efccb6f7-5641-4e0e-bd10-b4976e1bf68e'

    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $license.SkuId = $ENTERPRISEPACK
    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.AddLicenses = $license 
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses

    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $license.SkuId = $EMS
    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.AddLicenses = $license 
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses
}
function Add-VisioLicense($UserPrincipalName) {

    $VISIOCLIENT = 'c5928f49-12ba-48f7-ada3-0d743a3601d5'

    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $license.SkuId = $VISIOCLIENT
    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.AddLicenses = $license 
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses
}
function Add-ProjectLicense($UserPrincipalName) {

    $PROJECTPROFESSIONAL = '53818b1b-4a27-454b-8896-0dba576410e6'

    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $license.SkuId = $PROJECTPROFESSIONAL
    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.AddLicenses = $license 
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses
}

function Remove-O365License {

    $ENTERPRISEPACK = '6fd2c87f-b296-42f0-b197-1e91e994b900'
    $EMS = 'efccb6f7-5641-4e0e-bd10-b4976e1bf68e'

    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.RemoveLicenses = $ENTERPRISEPACK
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses

    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.RemoveLicenses = $EMS
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses
}

function Remove-VisioLicense {

    $VISIOCLIENT = 'c5928f49-12ba-48f7-ada3-0d743a3601d5'

    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.RemoveLicenses = $VISIOCLIENT
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses
}
function Remove-ProjectLicense {

    $PROJECTPROFESSIONAL = '53818b1b-4a27-454b-8896-0dba576410e6'

    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $AssignedLicenses.RemoveLicenses = $PROJECTPROFESSIONAL
    Set-AzureADUserLicense -ObjectId $UserPrincipalName -AssignedLicenses $AssignedLicenses
}
function Get-AssignedLicenses {
    Get-AzureADUserLicenseDetail -ObjectId $UserPrincipalName | Select SkuPartNumber
}
