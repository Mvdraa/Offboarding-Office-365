<#
  Script does the following:
  Converts Mailbox to Shared
  Blocks sign in
  Appends "Archief - "to display name
  Removes User from GAL
  Removes all assigned licenses
#>

<# Required: 
  Microsoft.Graph.Users.Actions
  Microsoft.Graph.Identity.DirectoryManagement
  Exchange Online Management
#>

#Forces shell to check for uninitialized variables.
Set-Strictmode -version 2

#Users.Actions
if (Get-Module -ListAvailable -name Microsoft.Graph.Users.Actions) {
  write-host "Microsoft.Graph.Users.Actions installed! Req (1/3)" -ForegroundColor Green
}
else {
  Install-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
  write-host "Module just got installed! Req (1/3)" -ForegroundColor Green
}

#Identity.DirectoryManagement
if (Get-Module -ListAvailable -name Microsoft.Graph.Identity.DirectoryManagement) {
  write-host "Microsoft.Graph.Identity.DirectoryManagement installed! Req (2/3)" -ForegroundColor Green
}
else {
  Install-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
  write-host "Module just got installed! Req (2/3)" -ForegroundColor Green
} 

#ExchangeOnlineManagement
if (Get-Module -ListAvailable -name ExchangeOnlineManagement) {
  write-host "ExchangeOnlineManagement installed! (req 3/3)" -ForegroundColor Green
}
else {
  Install-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
  write-host "Module just got installed! Req (3/3)" -ForegroundColor Green
} 


#Set user to Offboard
Connect-ExchangeOnline
$user = Read-Host ("Vul e-mailadres in")

#Hide from GAL
if ((get-mailbox $user).HiddenFromAddressListsEnabled)  {
  write-host "User already hidden from GAL" -ForegroundColor Green
} else {
  Set-Mailbox $user -HiddenFromAddressListsEnabled $true
}

#Convert Mailbox to shared.
if ((get-mailbox $user).RecipientTypeDetails -eq "UserMailbox"){
  Set-Mailbox $user -type Shared
} else {
  Write-Host "User is not an User mailbox" -ForegroundColor Red
  Pause
  Exit
}

#Connect to Microsoft Graph for license cmdlets.
Connect-Graph User.ReadWrite.All, Organization.Read.All

#Set user to MgUser
$mguser = Get-MgUser -UserId $user 

#Block sign-in,change display name
Update-MgUser -UserId $user -DisplayName ("Archief - " + $mguser.displayname) -AccountEnabled:$false


#Get assigned license information for user
$lic = Get-MgUserLicenseDetail -userid $user


<#Import CSV File with all license results from Get-MgUserLicenseDetails, List notation: !SEMICOLON DELIMITED ONLY
  Get CSV File from: 
  https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
#>

$licList = import-csv ".\Product names and service plan identifiers for licensing.csv" -Delimiter ';' | Where-Object {$lic.SkuID -match $_.GUID} | Select-Object Product_Display_Name -unique


#Removes License

If ($null -eq (Get-MgUserLicenseDetail -UserId $user)){
  Write-Host "No licenses currently assigned to $user" -ForegroundColor Red
} else {
  write-host "Removed the following licenses:" -ForegroundColor White -BackgroundColor Green
  $licList
  Set-MgUserLicense -UserId $user -RemoveLicenses $lic.SkuId -AddLicenses @{}
}

#Pause to read which licenses got removed
#Disconnect current sessions to prevent too many active connection errors
Read-Host -Prompt "User archived Press enter to exit"
Disconnect-Graph
Disconnect-ExchangeOnline -Confirm:$false
