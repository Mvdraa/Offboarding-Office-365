# Offboarding-Office-365
Offboarding using Microsoft.Graph/ExchangeOnlineManagement

Requirements: <br>
Install-Module Microsoft.Graph <br>
Download csv file with all license details in same location as the script from: <br>
https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference


Update 29-11-2022
Offboarding based on CSV file
Add e-mail to License 
Converts to shared mailbox, appends text to Shared Mailbox, hides from GAL, removes all licenses on mailbox.

TO-DO:
Keep track of removed licenses
