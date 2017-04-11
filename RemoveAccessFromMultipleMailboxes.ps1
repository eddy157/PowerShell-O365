####################################################### 
# Remove Mailbox Access
# Script by Eddy 
# Version 1.0
# 4/10/2017
# Note: If import Fails Copy the Data and Create a new .csv file
####################################################### 

# Connect to Office 365
Write-Host "Connecting to Office 365" -foregroundcolor green -backgroundcolor black
Start-Sleep -s 2
$credential = Get-Credential
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication "Basic" -AllowRedirection
Import-PSSession $exchangeSession -DisableNameChecking

####################################################### 
#Color Functions
function Receive-OutputG 
{
 process { Write-Host $_ -ForegroundColor Green }
} 

function Receive-OutputR 
{
 process { Write-Host $_ -ForegroundColor Red }
} 

function Receive-OutputY 
{
 process { Write-Host $_ -ForegroundColor Yellow }
} 

#Variables
$startDTM = (Get-Date)

#CSV and Log File Location
$path = "C:\"
$FPath = "C:\ExportData.csv"
$logfile = $path + "\logfile.txt"

#Clean-up logs from previous run; create new logs and timestamp them
Remove-Item $logfile
(get-date).DateTime | add-content -path $logfile

#Get User Name for O365
$O365User = Read-Host "Provide the User Name for O365 in this format:UserUPNLogon@kidsii.com - anakin@kidsii.com"

####################################################### 
# Removing Access to Mailbox
####################################################### 

Write-Output "Removing Access to Mailboxes" | add-content $logfile -passthru | Receive-OutputY
Write-Output "Connected to O365" | add-content $logfile -passthru | Receive-OutputY

#Importing CSV File
Import-Csv $FPath | ForEach-Object {
 $Mailbox = $_."EMAIL ADDRESS"

#Removing Access to Mailbox
Write-Output "Looking for Access on: $Mailbox" | add-content $logfile -passthru | Receive-OutputY
$UserMailbox=Get-MailboxPermission $Mailbox | Where { ($_.User -like $O365User) } | select -expandProperty User
If ($O365User -eq $UserMailbox)
{
 Remove-MailboxPermission -Identity $Mailbox -user $O365User -AccessRights FullAccess -InheritanceType All -Confirm:$false
 Write-Output "Full Access Found, Removing" | add-content $logfile -passthru | Receive-OutputG
}
else
{
  Write-Output "No Access Found" | add-content $logfile -passthru | Receive-OutputR
}

# Removing SendAs
$UserSendAs = Get-RecipientPermission $Mailbox | Where { ($_.Trustee -like $O365User) } | select -expandProperty Trustee
If ($O365User -eq $UserSendAs)
{
 Remove-RecipientPermission $Mailbox -AccessRights SendAs -Trustee $O365User -Confirm:$false
 Write-Output "Send As Access Found, Removing..." | add-content $logfile -passthru | Receive-OutputG
}
else
{
 Write-Output "Send As Access not Found" | add-content $logfile -passthru | Receive-OutputR
}

# Removing Send On Behalf
 Set-Mailbox $Mailbox -GrantSendOnBehalfTo @{remove=$O365User}
 Write-Output "All Access for user: $O365User has been removed from Mailbox: $Mailbox" | add-content $logfile -passthru | Receive-OutputY

}

# Discconect from O365
$Osession = Get-PSSession | select -expandProperty Name
Remove-PSSession -Name $Osession
Write-Output "Discconected from O365" | add-content $logfile -passthru | Receive-OutputY

# End
#######################################################