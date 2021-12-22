#Check for MSOL module

if ($null -eq (Get-Module -ListAvailable -Name MSonline)) {
    Write-Host "ERROR: MSonline module required to run this script!" -ForegroundColor Red
    Read-Host "To install, run 'Install-Module MSOnline'. Press enter to exit"
} else {
    Import-Module MSOnline
}

$TenantInfo = Get-MsolCompanyInformation

Function Select-FolderDialog {
    param([string]$Description="Select Folder",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

$Users = Get-MsolUser -EnabledFilter EnabledOnly | Where-Object {$_.IsLicensed -eq $True} | Sort-Object UserPrincipalName
$CsvOutput = foreach ($MsolUser in $Users){
    $MFAMethod = $MsolUser.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true} | Select-Object -ExpandProperty MethodType
    If (($MsolUser.StrongAuthenticationRequirements) -or ($MsolUser.StrongAuthenticationMethods)) {
       Switch ($MFAMethod) {
        "OneWaySMS" { $Method = "SMS token" }
        "TwoWayVoiceMobile" { $Method = "Phone call verification" }
        "PhoneAppOTP" { $Method = "Hardware token or authenticator app" }
        "PhoneAppNotification" { $Method = "Authenticator app" }
        }
        #Write-Host "$($MsolUser.DisplayName)'s MFA method is $Method"
    } else {
        #Write-Host "$($MsolUser.DisplayName) does not have an MFA method configured"
        $Method = "MFA not configured"
    }
    New-Object -TypeName PSCustomObject -Property @{
        Name = $MsolUser.DisplayName
        Method = $Method
    } | Select-Object Name,Method
}

Read-Host "Data gathered! Press enter to select where you want the CSV"

$Location = Select-FolderDialog

$CsvOutput | Export-Csv "$Location\$($TenantInfo.DisplayName) MFA.csv" -NoTypeInformation