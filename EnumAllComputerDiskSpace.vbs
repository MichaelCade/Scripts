'#------        EnumAllComputerDiskSpace.vbs            ------
'#------                                                ------
'# author:          Ed Morgan [ed.morgan@hp.com]        ------
'# version:         0.2 - 02/08/2011                    ------
'# changelog:       0.2 - Logging/version control       ------
'#------------------------------------------------------------

option explicit
Wscript.Echo "Started..."
' get the hostname, strip first four chars to get site ID
Dim WshNetwork
Set WshNetwork = WScript.CreateObject("WScript.Network")

Dim HostName
HostName = WshNetwork.ComputerName

Dim SiteID 
SiteID = Left(HostName, 4)

Dim strDate, strMonth, strYear, strHours, strMins
strDate = Day(now)
strMonth = Month (now)
strYear = Year (now)
strHours = Hour (time)
strMins = Minute (time)

Dim LogDate
LogDate = ""& strYear &"_" & strMonth &"_" &strDate &"_" & strHours & "_" & strMins &""

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strDirectory, strFile, objFile, objFolder
strDirectory = "."
strFile = "\" & SiteID & "_" & LogDate & "_DiskSpace.log"
Set objFile = objFSO.CreateTextFile(strDirectory & strFile)

set objFile = nothing
set objFolder = Nothing

Const ForAppending = 8
Dim DebugLog
Set DebugLog = objFSO.OpenTextFile _
(strDirectory & strFile, ForAppending, True)


DebugLog.WriteBlankLines 3
DebugLog.WriteLine("-- " & now & " --- Logging Started ---") 
DebugLog.WriteBlankLines 1


Const ADS_SCOPE_SUBTREE = 2
Dim adsRootDSE
Set adsRootDSE = GetObject("LDAP://rootDSE")
Dim strDomainDN, objConnection, objCommand, objRecordSet

strDomainDN = adsRootDSE.Get("defaultNamingContext")

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = "Select Name from 'LDAP://" & strDomainDN & "'Where objectClass='computer'and Name='" & SiteID & "*' order by name"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst


Const HARD_DISK = 3

on error resume next

Do Until objRecordSet.EOF

DebugLog.WriteLine objRecordSet.Fields("Name").Value
			
DebugLog.WriteLine

Dim strComputer, objDiskItem, objDISKWMIService, colDiskItems

strComputer = objRecordSet.Fields("Name").Value

set objDiskItem = nothing
Set objDISKWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colDiskItems = objDISKWMIService.ExecQuery("Select * from Win32_Volume")

For Each objDiskItem in colDiskItems

    if objDiskItem.DriveType = 3 then
	    If objDiskItem.Label <> Null Then
		    DebugLog.WriteLine objDiskItem.Name & space(30-Len(objDiskItem.Name)) & " " & objDiskItem.Label + space(14-Len(objDiskItem.Label)) & "Capacity=" & space(6-Len(FormatNumber(objDiskItem.Capacity /1073741824))) & FormatNumber(objDiskItem.Capacity /1073741824) & "GB, " & space(6-Len(FormatNumber(objDiskItem.FreeSpace/1073741824))) & FormatNumber(objDiskItem.FreeSpace/1073741824) & "GB free"
	    Else
		    DebugLog.WriteLine objDiskItem.Name & space(30-Len(objDiskItem.Name)) & " " & space(14) & "Capacity=" & space(6-Len(FormatNumber(objDiskItem.Capacity /1073741824))) & FormatNumber(objDiskItem.Capacity /1073741824) & "GB, " & space(6-Len(FormatNumber(objDiskItem.FreeSpace/1073741824))) & FormatNumber(objDiskItem.FreeSpace/1073741824) & "GB free"
        End If
  	end if
Next

DebugLog.WriteLine
objRecordSet.MoveNext

Loop


DebugLog.WriteLine("-- " & now & " --- Getting DB Sizes ---") 

Dim strDBServerName, strDBName, objSQLServer, objDB

strDBServerName = siteID & "DBServer1\Apps"
strDBName = "SX"


Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName 

Set objDB = objSQLServer.Databases(strDBName)

DebugLog.WriteLine "Total size of Sanctuary " & strDBName & " DB is " & objDB.Size & "Mb"
objSQLServer.Close



strDBServerName = SiteID & "DBServer2\NETCOOL"
strDBName = "Reporter"

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName 

Set objDB = objSQLServer.Databases(strDBName)

DebugLog.WriteLine "Total size of Netcool " & strDBName & " DB is " & objDB.Size & "Mb"
objSQLServer.Close


strDBServerName = SiteID & "DBServer2\NETCOOL"
strDBName = "ReporterNFSM"

Set objSQLServer = CreateObject("SQLDMO.SQLServer")
objSQLServer.LoginSecure = True
objSQLServer.Connect strDBServerName 

Set objDB = objSQLServer.Databases(strDBName)

DebugLog.WriteLine "Total size of Netcool " & strDBName & " DB is " & objDB.Size & "Mb"
objSQLServer.Close

Wscript.Echo "Finished"
DebugLog.WriteLine("-- " & now & " --- Logging Finished ---") 



			
