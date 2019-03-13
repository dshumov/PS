param([parameter(Mandatory=$true)] [String]$FileCSV)
$listOU=Import-CSV $FileCSV -Delimiter ";"


ForEach($OU in $listOU){

	try{
	#Get Name and Path from the source file
	$OUName = $OU.Name
	$OUPath = $OU.Path

	#Display the name and path of the new OU
	Write-Host -Foregroundcolor Yellow $OUName $OUPath
	
	#Create OU
	New-ADOrganizationalUnit -Name "$OUName" -Path "$OUPath"
	
	#Display confirmation
	Write-Host -ForegroundColor Green "OU $OUName created"
	}catch{
	
		Write-Host $error[0].Exception.Message
	}

}