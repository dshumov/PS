<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2016 v5.2.129
	 Created on:   	06.03.2019 16:52
	 Created by:   	IT
	 Organization: 	CraftGroup
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>
Import-Module ActiveDirectory

$Dir = "\\WSFS-01\Shares$\DOCS_users"
$Users = Get-ADUser -Filter * -SearchBase "OU=New base,DC=group,DC=com"
foreach ($User in $Users)
{
	$User1 = $User.sAMAccountName
	$Path = New-Item -ItemType Directory -Path $Dir -Name $User1
	$User = $User.sAMAccountName
	$Path1 = $Dir + "\" + $User1
	$Args = New-Object  system.security.accesscontrol.filesystemaccessrule($User, "FullControl, Synchronize", "ContainerInherit, ObjectInherit", "None", "Allow")
	$ACL = Get-Acl $Path
	$ACL.SetAccessRule($Args)
	Set-Acl -Path $Path -AclObject $ACL
}