
cls


#gwmi -Class Win32_LogicalDisk -Filter "DriveType='3'" | Where-Object {$_.Size -ne $null} | Select-Object DeviceID |
#ForEach-Object {Get-ChildItem $_.DeviceID -Include *.txt -Recurse | ForEach-Object {Copy-Item $_.FullName G:\Temp -Recurse }}

#gwmi -Class Win32_LogicalDisk -ComputerName cp555

#Get-Item E:\1 | ForEach-Object {Copy-Item $_.FullName G:\Temp -Recurse -Exclude *карта*  }

#Copy-Item E:\1\ -Destination G:\Temp -Filter *.txt -Exclude  -Recurse 

$mass = Get-ChildItem E:\1\ -Include *.txt -Recurse -Exclude "*Карта*" 

if ($mass[6].DirectoryName -notcontains "4324")
{
$DirectoryName = $mass[6].DirectoryName -split ":", 0, "simplematch"
$DirectoryName = $DirectoryName[1..1]
$DirectoryName
$d = New-Item G:\Temp\$DirectoryName -ItemType directory
Copy-Item $mass[4].FullName $d -Recurse 
}
else
{
$d = "Fuck"
$d
}