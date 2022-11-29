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
#import csv file with multiple users
#1 row based on email, leave column 1 with e-mail
$users = import-csv .\OffboardThese.csv

#Start looping through users in imported CSV
foreach ($user in $users) {
  #Skip if Shared Mailbox already
  if((get-Mailbox $user.email).RecipientTypeDetails -ne "UserMailbox") {
    Write-Host "User is not a user mailbox, moving on to next user" -ForegroundColor red
  } 
  else {
    #Hide from GAL
    if ((get-Mailbox $user.email).HiddenFromAddressListsEnabled)  {
      write-host "User already hidden from GAL" -ForegroundColor Green
    } else {
      Set-Mailbox $user.email -HiddenFromAddressListsEnabled $true
    }

    #Convert Mailbox to shared.
    if ((get-Mailbox $user.email).RecipientTypeDetails -eq "UserMailbox"){
      Set-Mailbox $user.email -type Shared
    } else {
      Write-Host "User is not an User mailbox" -ForegroundColor Red
      Pause
      Exit
    }

    #Connect to Microsoft Graph for license cmdlets.
    Connect-Graph User.ReadWrite.All, Organization.Read.All


    #Set user to MgUser
    $mguser = Get-MgUser -UserId $user.email 

    #Block sign-in,change display name
    Update-MgUser -UserId $user.email -DisplayName ("Archief - " + $mguser.displayname) -AccountEnabled:$false


    #Get assigned license information for user
    $lic = Get-MgUserLicenseDetail -userid $user.email
    #Removes License

    If ($null -eq (Get-MgUserLicenseDetail -UserId $user.email)){
      Write-Host "No licenses currently assigned to $user.email" -ForegroundColor Red
    } else {
      Set-MgUserLicense -UserId $user.email -RemoveLicenses $lic.SkuId -AddLicenses @{}
    }
  }
}
#Pause to read which licenses got removed
#Disconnect current sessions to prevent too many active connection errors
Read-Host -Prompt "User archived Press enter to exit"
Disconnect-Graph
Disconnect-ExchangeOnline -Confirm:$false