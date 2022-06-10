Import-Module ActiveDirectory

$Username = Read-Host -Prompt "Enter username of mover"
$Title = Read-Host -Prompt "Enter new job title"
$OU = "OU=Email in O365,OU=SelwoodUsers,OU=FLIP,DC=selwoodhousing,DC=local"
$Manager = Read-Host -Prompt "Enter Managers Username"

IF (@(Get-ADUser -Filter * -Properties title, department -SearchBase $OU -SearchScope Subtree | where {$_.title -eq $Title}).Count -gt 0 -and (@(Get-ADUser -Filter * -Properties title, department -SearchBase $OU -SearchScope Subtree | where {$_.title -eq $Title}).Count -le 7) ){
    $Equiv = Get-ADUser -Filter * -Properties title, department -SearchBase $OU -SearchScope Subtree | where {$_.title -eq $Title} | Select-Object -First 1
    }

    Else {
    Get-ADUser -Filter * -Properties Displayname, Title, Department -SearchBase $OU -SearchScope Subtree | select Displayname, Title, Department | where {$_.title -eq $Title} | ft
    
    $NameMatch = Read-Host -Prompt "Please choose an entry from the list in the format Firstname.Surname"
    $Equiv = Get-ADUser -Identity $NameMatch
    }

Write-Host "Equivalent user is" $Equiv.Name -ForegroundColor Cyan

try{
    Set-ADUser -Identity $Username -Title $Title -Description $Title

    Get-ADUser -Identity $Username -Properties name, memberof | Select-Object memberof -ExpandProperty memberof | Remove-ADGroupMember -Members $Username -Confirm:$false

    Get-ADUser -Identity $Equiv -Properties name, memberof | Select-Object memberof -ExpandProperty memberof | Add-ADGroupMember -Members $Username

    $NewUser = Get-ADUser -Identity $UserName
    $ManagerDN = Get-ADUser -Identity $Manager
    $NewUser | Set-ADUser -Manager $ManagerDN

    $Department = Get-ADUser -Identity $Equiv -Properties name, department | Select-Object department
    Set-ADUser -Identity $Username -Department $Department.department
    Write-Host "Successfully updated user account to match" $equiv.Name -ForegroundColor Green
    }

Catch{
    Write-Host "Failed to copy equivalent user" -ForegroundColor Red
    }