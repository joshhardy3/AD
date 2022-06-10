<#
UPDATE NOTES:

08/06/2022:
    Removed need to manually add file path of restore doc, will auto search ADDC for file name
    Imports 2 docs, one for memberof and one for AD Details
    Removed Auto reply
    Converts mailbox to normal from shared
    Added silently continue for Memberof groups

#>
$adminusername = Get-Content "C:\Temp\email.txt"
$Adminpass = Get-Content "C:\Temp\pword.txt" | ConvertTo-SecureString
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $adminusername, $Adminpass


$Name = Read-Host -Prompt "Enter Username"

$Username = Get-ADUser -Identity $Name


$ADDetails = "\\vs-addc-l-01\SCRIPTS\Scriptdocs\AD_Memberof_Backup\" + $Name + "ADDetails.csv"
$NewGroups = "\\vs-addc-l-01\SCRIPTS\Scriptdocs\AD_Memberof_Backup\" + $Name + "MemberOf.csv"

$Details = Import-Csv -Path $ADDetails

Try{

    $Groups = import-csv -Path $NewGroups
    Foreach ($Group in $Groups){
        Get-ADGroup -Filter * -Properties Distinguishedname | Where {$_.distinguishedname -eq $Group.distinguishedName} | Add-ADGroupMember -Members $Username -ErrorAction SilentlyContinue
        }
    Write-Host "Restored Memberof Groups" -ForegroundColor Cyan
    }

Catch{
    Write-Host "Failed to Restore Memberof Groups"
    }

Try{
    Get-ADUser -Identity $Name | Set-ADUser -Description $Details.Title -Manager $Details.Manager -Title $Details.Title -Department $Details.Department -Company "Selwood Housing" -Enabled $true
    
    Get-ADUser -Identity $Name | Set-ADUser -Add @{ physicalDeliveryOfficeName = "Bryer Ash"}
    Write-Host "Restored AD Details" -ForegroundColor Cyan
    }

Catch{
    Write-Host "Unable to restore AD user Details"
    }

Try{
    
    Get-ADUser -Identity $UserName | Move-ADObject -TargetPath "OU=Email in O365,OU=SelwoodUsers,OU=FLIP,DC=selwoodhousing,DC=local" -Confirm:$false
    Write-Host "Moved User to 365 OU" -ForegroundColor Cyan
    }

Catch{
    Write-Host "Unable to move to 365 OU"
    }

Try{
    
    Invoke-Command -ComputerName vs-azac-l-01 -ScriptBlock {Start-ADSYNCSYNCCYCLE}
    Write-Host "Starting AD Sync" -ForegroundColor Cyan
    Start-Sleep -s 60
    }

Catch{
    Write-Host "Unable to start AD Sync"
    }

Try{
    
    Connect-ExchangeOnline -Credential $Cred

    Get-Mailbox -Identity $Name | Set-MailboxAutoReplyConfiguration -AutoReplyState Disabled
    Get-Mailbox -Identity $Name | Set-Mailbox -Type Regular
    
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Converted mailbox from Shared and disabled Auto Reply" -ForegroundColor Cyan
    }

Catch{
    
    Write-Host "Unable to convert mailbox"
    }

$Number = $Details.Telephonenumber
Write-Host $Number

Write-Host "Successfully restored AD account for $Name, please manually update number to $Number" -ForegroundColor Green
