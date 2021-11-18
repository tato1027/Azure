# Author: Anton Petrianik
# Copying a data from Department attribute to extensionAttribute11, and cut to 65 letters if lengh is more than 65.



$ou = "" # Distinguished name of CN with users
$users = Get-ADUser -Filter * -SearchBase $ou -Properties Department, extensionAttribute11 | Select-Object SamAccountName, Department, extensionAttribute11

function Copy-DepartmentAttribute {
    param($user)
        
    if ($user.Department -ne $null -and $user.Department.Length -gt 64) {
        $department = $user.Department.Substring(0,64)
    }
    elseif($user.Department -ne $null -and $user.Department.Length -le 64) {
        $department = $user.Department    
    }  
       
    $thisUser = Get-ADUser -Identity $user.SamAccountName -Properties extensionAttribute11
    Set-ADUser –Identity $thisUser -add @{"extensionattribute11"=$department.ToString()} -Verbose
}

foreach ($user in $users) {
    Copy-DepartmentAttribute -user $user
}
