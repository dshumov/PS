Dim adsinfo, ThisComp, oUser

' Определяем объекты
Set adsinfo = CreateObject("adsysteminfo")
Set ThisComp = GetObject("LDAP://" & adsinfo.ComputerName)
Set oUser = GetObject("LDAP://" & adsinfo.UserName)

' Заносим данные в AD 
' В поле description компьютера пишем имя пользователя, и время входа
Thiscomp.put "description", "Logged off: " + oUser.cn + " " + CStr(Now)
ThisComp.Setinfo

'В поле info пользователя пишем имя компьютера и время входа
oUser.put "info", "Logged off: " + ThisComp.cn + " " + CStr(Now)
oUser.Setinfo 
wscript.quit