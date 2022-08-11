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

# Check for Modules, if not available install them. 
#Exit script if can't install

#Microsoft.Grap
if (Get-Module -ListAvailable -name Microsoft.Graph) {
  write-host "Microsoft.Graph installed! Req (1/4)" -ForegroundColor Green
}
else {
  try {
    Install-Module Microsoft.Graph -Erroraction Stop
    write-host "Module just got installed! Req (1/4)" -ForegroundColor Green
  } 
  catch {
    write-host "Could not install required graph module (Microsoft.Graph)" -ForegroundColor Red
    Pause
    exit
  } 
}

#Users.Actions
if (Get-Module -ListAvailable -name Microsoft.Graph.Users.Actions) {
  write-host "Microsoft.Graph.Users.Actions installed! Req (2/4)" -ForegroundColor Green
}
else {
  try {
    Install-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    write-host "Module just got installed! Req (2/4)" -ForegroundColor Green
  } 
  catch {
    write-host "Could not install required graph module (Microsoft.Graph.Users.Action)" -ForegroundColor Red
    Pause
    exit
  } 
}

#Identity.DirectoryManagement
if (Get-Module -ListAvailable -name Microsoft.Graph.Identity.DirectoryManagement) {
  write-host "Microsoft.Graph.Identity.DirectoryManagement installed! Req (3/4)" -ForegroundColor Green
}
else {
  try {
    Install-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    write-host "Module just got installed! Req (3/4)" -ForegroundColor Green
  } 
  catch {
    write-host "Could not install required graph module (Microsoft.Graph.Identity.DirectoryManagement)" -ForegroundColor Red
    Pause
    exit
  }
}

#ExchangeOnlineManagement
if (Get-Module -ListAvailable -name ExchangeOnlineManagement) {
  write-host "ExchangeOnlineManagement installed! (req 4/4)" -ForegroundColor Green
}
else {
  try {
    Install-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
    write-host "Module just got installed! Req (4/4)" -ForegroundColor Green
  } 
  catch {
    write-host "Could not install required graph module (ExchangeOnlineManagement)" -ForegroundColor Red
    Pause
    exit
  }
}

#Set user to Offboard
Connect-ExchangeOnline
$user = Read-Host ("Vul e-mailadres in")

#Hide from GAL
if ((get-mailbox $user).HiddenFromAddressListsEnabled)  {
  write-host "User already hidden from GAL" -ForegroundColor Green
  continue
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


#Block sing-in,change display name
Update-MgUser -UserId $user -DisplayName ("Archief - " + $mguser.displayname) -AccountEnabled:$false

#Removes License
$licId = Get-MgUserLicenseDetail -userid $user
Set-MgUserLicense -UserId $user -RemoveLicenses $licId.SkuId -AddLicenses @{}


