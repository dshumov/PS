$user = 'Domain\user'
$folders = "Имя папки с подпапками для работы"

##Удаление наследования в дочерних папках
Foreach($folder in $folders) { 
 icacls $folder /inheritance:d
 Get-ChildItem -Path $folder -Recurse | ?{$_.PSisContainer} | foreach {$subfolder = $_.FullName; icacls $subfolder /inheritance:d}
}

##Удаление группы безопасности в дочерних папках
Foreach($folder in $folders) { 
 Get-ChildItem -Path $folder -Recurse | ?{$_.PSisContainer} | foreach {$subfolder = $_.FullName; icacls $subfolder /remove "GROUP\Сервис файловой сервер" /T}
}

##Удаление пользователя с дочерних папок
Foreach($folder in $folders) { 
Get-ChildItem -Path $folder -Recurse | Foreach {$object = $_.FullName; 
 $acls = Get-Acl -Path $_.FullName; 
 foreach ($acl in $acls) {
  foreach ($access in $acl.Access) {
   foreach ($value in $access.IdentityReference.Value) {
    if ($value -eq $user) {
     $acl.RemoveAccessRule($access)
} } } }
Set-Acl -path $object -aclObject $acl
} }
##Включение наследования на дочерних папках
Foreach($folder in $folders) { 
Get-ChildItem -Path $folder -Recurse | ?{$_.PSisContainer} | foreach {$subfolder = $_.FullName; icacls $subfolder /inheritance:e}
}