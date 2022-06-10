<#
UPDATE NOTES:

07/06/2022:
    Amended AD backup to create 2 separate files for Memberof and AD Details

08/06/2022:
    Removed confirmation on updating memberof, making script require no input

#>

#Connect AD
import-module activedirectory
Import-Module MSOnline

#Params
$FirstName = Read-Host -Prompt "Enter the Leavers First Name"
$Surname = Read-Host -Prompt "Enter the Leavers Surname"
$AcctoMbox = Read-Host -Prompt "Does anyone require access to the Leavers mailbox? If so, enter their Name, if not leave blank"
$UserName = $FirstName + "." + $Surname
$Email = $UserName + "@Domain.com"
$Date = Get-Date -Format "dd/MM/yyyy"
$Description = "Leaver - " + $Date
$UPN = $UserName + "@Domain.com"
$Name = $FirstName + " " + $Surname
$AutoReply = "Thank you for your email, however " + $Name + " no longer works for "". Please call "" and a member of our Customer Service team will be able to provide you with information on where to send the information too."
$adminusername = Get-Content "C:\Temp\Email.txt"
$pass = Get-Content "C:\Temp\pword.txt" | ConvertTo-SecureString
$Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminusername, $pass
$Date2 = Get-Date -Format "dd.MM.yyyy"
$BCKPath = "FilePath" + $UserName + "MemberOf.csv"
$LogFile = "FilePath" + $username + "ADDetails.csv"




Function ADBackup{

    Try{
    Get-ADUser -Identity $username -Properties Name, Title, Manager, telephonenumber, Department, Office | select Name, Title, Manager, Telephonenumber, Department, Office | Export-Csv -Path $LogFile -NoTypeInformation
    Get-ADPrincipalGroupMembership -Identity $UserName | Export-Csv $BCKPath -NoTypeInformation
    Write-Host "Created Memberof backup file" + $BCKPath -ForegroundColor Cyan
    }

    Catch{

    Write-Host "Failed to create AD Backup"
    }
}

Function ADCleanUp{

    Try{

    Set-ADUser -identity $UserName -Description $Description -enabled $False -Confirm:$False

    Set-ADUser -Identity $UserName -Clear Manager -Confirm:$False
    Set-ADUser -Identity $UserName -Clear Telephonenumber -Confirm:$False
    Set-ADUser -Identity $UserName -Clear Mobile -Confirm:$False
    
    Get-ADUser -Identity $UserName -Properties MemberOf | ForEach-Object {
    $_.MemberOf | Remove-ADGroupMember -Members $UserName -Confirm:$False
        }
    Write-Host "Successfully updated AD details" -ForegroundColor Cyan
    }

    Catch{

    Write-Host "Unable to update AD Details"
    }
}
    

#Not Active as shouldnt be needed, licenses pull from AD group and will be removed when added to Leavers
Function RemoveLicenses{


    Import-Module MSOnline
    Start-Sleep -Seconds 5
    Connect-MsolService -Credential $creds
    
    
    Try{
    
    (get-MsolUser -UserPrincipalName $UPN).licenses.AccountSkuId | foreach{
    Set-MsolUserLicense -UserPrincipalName $UPN -RemoveLicenses $_
        }
    Write-Host "Successfully Removed Licenses" -ForegroundColor Cyan
    }

    Catch{

    Write-Host "Error removing licenses, please remove manually"
    }
}


Function ConnectExo{

    Try{

    Connect-ExchangeOnline -Credential $creds
    Write-Host "Connected to Exchange Online, starting 10 second pause" -ForegroundColor Cyan
    }


    Catch{

    Write-Host "Unable to connect to Exchange online"
    }

    finally{

    Start-Sleep -Seconds 10
    }
}


Function AddAutoReplyandConvert{

    Try{

    Set-MailboxAutoReplyConfiguration -Identity $Email -AutoReplyState Enabled -InternalMessage $AutoReply -ExternalMessage $AutoReply
    Write-Host "Successfully set Auto Reply, starting 10 second pause" -ForegroundColor Cyan
    }

    Catch{

    Write-Host "Failed to set auto reply"
    }

Start-Sleep -Seconds 10


    Try{

    Set-mailbox -identity $Email -type Shared
    Write-Host "Converted mailbox to shared, starting 10 second pause" -ForegroundColor Cyan
    }

    Catch{

    Write-Host "Failed to convert to shared"
    }


Start-Sleep -Seconds 30
}

function MBoxAccess{

    Try{
    $User = Get-ADUser -Identity $AcctoMbox

    foreach ($User in $Users){
        Add-MailboxFolderPermission -Identity $UserName -User $user.SamAccountName -AccessRights Fullaccess
        }
    }

    Catch{
    Write-Host "No User to be added to full access" -ForegroundColor Cyan
    }

Disconnect-ExchangeOnline -Confirm:$False
}

Function MovetoLeavers{

    Try{ 
    
    Get-ADUser -Identity $UserName | Move-ADObject -TargetPath "Leaver OU" -Confirm:$false
    Write-host "Moved $Firstname $Surname To Leavers OU" -ForegroundColor Cyan
    }

    Catch{
    Write-Host "Unable to move user to Leavers OU"
    }
}


ADBackup
ADCleanUp
ConnectExo
AddAutoReplyandConvert
MBoxAccess
MovetoLeavers
Write-Host "Leaver process complete, please amend OneDrive access if needed -ForegroundColor Green
