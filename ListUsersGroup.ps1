# Author: Anton Petrianik
# Get users from group and get group memebership with regexp name

$users = Get-ADGroupMember -Identity ‘groupname’ -Recursive | select SamAccountName, Name 
$list = @()

foreach ($user in $users) {
    $groups = Get-ADPrincipalGroupMembership $user.SamAccountName | Where-Object {$_.name -match "*"} | Select-Object name
    $list += $user.SamAccountName + "," + $user.Name + "," + $groups.name
}

$list | Out-File -FilePath C:\temp\file.csv 
