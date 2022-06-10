<# 
UPDATE NOTES:

30/05/2022:
    Added SMTP address
    Added RRA
    Changed Equiv search to include subfolders of 365 OU
    If more than 7 people have matching job title, will prompt for manual entry so's not to confuse multi trade teams etc
    If role doesnt exist, will prompt to choose from list, ignore this and manually enter user

04/06/2022:
    Changed script to create new remote mailbox rather than enable remote mailbox
    Now creates mailbox on 365 and on prem and automatically creates AD user, which is then amended later in the scripts

07/06/2022:
    Removed parameter to set Primary SMTP from create mailbox command as caused issues with Remote Routing, have added it after creating mailbox
    Disable Email Address Policy in order to set Primary SMTP

08/06/2022:
    Added variable to ask if 1st touch is required, if yes, added template email to DRS manager to clipboard to paste and send
    Updated so if more than 1 person has job title, will ask for Equiv manually, so's not to confuse departments, until I can add the below

09/06/2022:
    Added Progress bar to AD sync rather than the blind 60 second pause
    Updated Equiv search so if there is more than 1 department within users with that Title, it will prompt for manual input, if
    users have the same department, it will select first one
#>


Import-Module ActiveDirectory

start-sleep -seconds 2

#Inputs
$FirstName = Read-Host -Prompt "Enter First Name"
$Initial = $FirstName.Substring(0,1)
$Surname = Read-Host -Prompt "Enter Surname"
$Title = Read-Host -Prompt "Enter Title"
$Manager = Read-Host -Prompt "Enter Managers Username"

#Variables
$UserName = $FirstName + "." + $Surname
$Email = $UserName + "@Domain.com"
$FullName = $FirstName + " " + $Surname
$Path = "OU=OUT,OU=OU,OU=OU,DC=Domain,DC=local"
$Company = "Company"
$OU = "OU=OUT,OU=OU,OU=OU,DC=Domain,DC=local"
$DC = "Domain controller"
$DC1 = "Domain Controller 1"
$Prox = $UserName + "@domain.mail.onmicrosoft.com"
$adminusername = Get-Content "C:\Temp\email.txt"
$Adminpass = Get-Content "C:\Temp\pword.txt" | ConvertTo-SecureString
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminusername, $Adminpass
$Prim = "SMTP:" + $Email
$RRA = $Initial +"." + $Surname + "@domain.mail.onmicrosoft.com"
$RRA1 = $UserName + "@domain.mail.onmicrosoft.com"
$SMTP1 = $Initial +"." + $Surname + "@domain.com"
$1stTouch = Read-Host -Prompt "Does this user require 1st Touch? (Yes or No)"

#Generate Pword - https://community.spiceworks.com/topic/2130057-generate-random-password
$Password = "!@#$%^&*0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz".tochararray()

$Newpas = ($Password | Get-Random -count 8) -join ''

$SecPw = ConvertTo-SecureString -String $Newpas -AsPlainText -Force

#Get Equiv
Try{
    $PotentialEquiv = Get-ADUser -Filter * -Properties Displayname, Title, Department -SearchBase $OU -SearchScope Subtree | where {$_.title -eq $Title}

    $TempDepart = Get-ADUser -Filter * -Properties Displayname, Title, Department -SearchBase $OU -SearchScope Subtree | where {$_.title -eq $Title} | Select-Object -First 1 | Select-Object Department -ExpandProperty Department


#Loops each user in the above varibale and checks if department matches the first user

    foreach ($user in $PotentialEquiv){
        If (@(Get-ADUser -Identity $User -Properties Department | Select-Object Department -ExpandProperty Department) -eq $TempDepart)
            {$Prompt = "No"}
        
        Else{
             $Prompt = "Yes"}
          }

    If ($Prompt -eq "No"){
        $Equiv = Get-ADUser -Filter * -Properties title, department -SearchBase $OU -SearchScope Subtree | where {$_.title -eq $Title} | Select-Object -First 1
         }

    If ($Prompt -eq "Yes"){
        Get-ADUser -Filter * -Properties sAMaccountname, Title, Department -SearchBase $OU -SearchScope Subtree | select sAMaccountname, Title, Department | where {$_.title -eq $Title} | ft
    
        $NameMatch = Read-Host -Prompt "Please choose an entry from the list in the format Firstname.Surname"
        $Equiv = Get-ADUser -Identity $NameMatch
        }
 
Write-Host "Equivalent user is" $Equiv.Name -ForegroundColor Cyan

}

Catch{
    
    Write-Host "Unable to find Equivalent User"
    }

#Connecting to Exchange Server
Write-Host "Attempting connection to Exchange Server" -ForegroundColor Cyan
Try{
    
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://"ExchangeServer.domain".local/PowerShell/ -Authentication Kerberos -Credential $Cred
    Import-PSSession $Session -DisableNameChecking -AllowClobber
    Write-Host "Successfulyl connected PowerShell to Exchange Server" -ForegroundColor Cyan
    }

Catch{
    
    Write-Host "Failed to connect to Exchange Server" -ForegroundColor Red
    }


#Creating Mailbox
Try{

   New-RemoteMailbox -name $FullName `
        -FirstName $FirstName `
        -LastName $Surname `
        -DisplayName $FullName `
        -UserPrincipalName $Email `
        -Password $SecPw `
        -DomainController $DC1


   Write-Host "Created Mailbox" -foregroundcolor cyan
        Start-Sleep -s 60
    }

Catch{
   
   Write-Host "Failed to create Mailbox" -ForegroundColor Red
   }
   
Try{

   Set-RemoteMailbox -Identity $UserName -EmailAddressPolicyEnabled $false
   Write-Host "Disabled Email Address Policy, Starting 30 second pause before setting SMTP" -ForegroundColor Cyan
        Start-Sleep -Seconds 30

   Get-RemoteMailbox $UserName | Set-RemoteMailbox -PrimarySmtpAddress $Email
   Write-Host "Successfully Updated SMTP" -ForegroundColor Cyan
   }

Catch{
    Write-Host "Failed to update SMTP address" -ForegroundColor Red
   
   }

Remove-PSSession $Session

#AD Sync - https://wragg.io/using-write-progress-to-provide-feedback-in-powershell/
Try{
    Invoke-Command -ComputerName "AAD Sync Server" -ScriptBlock {Start-ADSYNCSYNCCYCLE}
    For ($i=60; $i -gt 1; $i–-) {
    Write-Progress -Activity "Running AD Sync" -SecondsRemaining $i
    Start-Sleep 1
    }
write-host "Completed AD Sync" -ForegroundColor Cyan
}

Catch{
    Write-Host "Unable to Start AD Sync" -ForegroundColor Red
}

#Amending AD user account
Try{

    Set-ADUser -Identity $UserName `
    -GivenName $FirstName `
    -Surname $Surname `
    -Description $Title `
    -Title $Title `
    -ChangePasswordAtLogon $true `
    -Office $Office `
    -Company $Company `
    -Enabled $true `
    -Server "Domain Controller"
        start-sleep -Seconds 10

    Get-ADUser -Identity $Equiv -Properties memberof | Select-Object -ExpandProperty Memberof | Add-ADGroupMember -Members $UserName
    
    $Departs = Get-ADUser -Identity $Equiv -Properties department | Select-Object department
    foreach ($Depart in $Departs)
    {
    Set-ADUser -Identity $UserName -Add @{Department=$Depart.Department}
    }


    $NewUser = Get-ADUser -Identity $UserName
    $ManagerDN = Get-ADUser -Identity $Manager
    $NewUser | Set-ADUser -Manager $ManagerDN


    #Set-ADUser $UserName -Add @{ProxyAddresses=$Prim}
    #Set-ADUser $UserName -Add @{ProxyAddresses=$RRA1}
    #Set-ADUser $UserName -Add @{ProxyAddresses=$Prox}
    Write-host "Updated AD User, starting 10 second pause" -ForegroundColor Cyan
        Start-Sleep -Seconds 10 
}
Catch{
    Write-Host "Unable to create AD User" -ForegroundColor Red
}

#Updating Inheritance to allow Skype access
Try{
    $user = get-aduser -Identity $UserName -properties ntsecuritydescriptor

    $user.ntsecuritydescriptor.SetAccessRuleProtection($true,$false)
    Write-Host "Successfully Updated AD Inheritance" -ForegroundColor Cyan
}
Catch{
    Write-Host "Unable to update AD Inheritance, please update manually" -ForegroundColor Red
}

#Move to 365 OU
Try{
    Move-ADObject -Identity $user.DistinguishedName `
    -TargetPath $OU `
    -Confirm:$false
    
    Write-Host "Successfully Moved AD Account to 365 Users OU" -ForegroundColor Cyan
    }

Catch{
    Write-Host "Unable to move AD account, please complete manually" -ForegroundColor Red
    }
    
$EmailDepart = Get-ADUser -Identity $UserName -Properties Department | Select-Object Department -ExpandProperty Department

$EmailtoDRSManager = "Hi "",
 
I’m currently processing this new starter and they require setup in DRS before I can setup their 1st Touch:
 
Name of Starter: $FullName
Team: $EmailDepart
Job Title: $Title
E-Mail Address: $Email
Manager: $Manager
 
Would you be able to give me their Operative ID once DRS has been set up for this user?
 
Kind regards,
"


If(
    $1stTouch -eq "Yes"){
        Set-Clipboard -Value $EmailtoDRSManager
        Write-host "Please paste clipboard in an email to the DRS Manager"
        }
Else{}

Set-ADAccountPassword -Identity $UserName -NewPassword $SecPw


#Finished
Write-host "The account $FullName Has been created with the password $Newpas" -ForegroundColor Green
