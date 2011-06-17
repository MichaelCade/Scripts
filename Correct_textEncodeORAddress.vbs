'#------	Correct_textEncodedORAddress.vbs		------
'#													------
'# author: 		Ed Morgan [ed.morgan@hp.com]    	------
'# version:		0.2 - 16/06/2011					------
'# changelog: 	0.2 - Added error checking			------
'# any business stuff removed before git push       ------
'#--------------------------------------------------------	

Option Explicit
On Error Resume Next

Const CANCEL_SCRIPT = 7
Const SCRIPT_VERSION  = "v0.2"
Const FOR_APPENDING = 8
Const ADS_PROPERTY_APPEND = 2
Const ADS_PROPERTY_DELETE = 4

Dim intResponse, intUserRoleCount, intMalformedAddresses, intWellformedAddresses, intCorrectedAddresses
Dim strLogFilePath, strScriptName, strTitle, strUserDN, strOldX400Addr, strNewX400Addr, strUserName
Dim objRootDSE, objConnection, objCommand, objRecordSet, objUser, objFSO, objLogFile

strScriptName = Split(WScript.ScriptName, ".")(0)
strTitle = strScriptName & " " & SCRIPT_VERSION

' Confirm we're running in the right domain.
intResponse = MsgBox("This should NOT be run in XXX domains." & vbCRLF & _
				"This will correct all User/Role X.400 addresses with XXX." & vbCRLF & _
                "All well-formed X.400 addresses will be left unchanged." & vbCRLF & vbCRLF & _
                "Do you want to Continue?", vbYesNo, strTitle & " All Users/Roles")
				
If intResponse = CANCEL_SCRIPT Then
	WScript.Echo "Command Cancelled."
	WScript.Quit(0)
End If 

' Create Log File
strLogFilePath = "./" & strScriptName & ".log"
Wscript.StdOut.WriteLine "Log File Name: " & strLogFilePath & vbCrLf

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Append Existing File
Set objLogFile = objFSO.OpenTextFile(strLogFilePath, FOR_APPENDING, True)

If Err.Number <> 0 then
	Msgbox "ERROR COULD NOT CREATE LOG FILE: " & UCase(strLogFilePath)
	WScript.Quit(1)
End If

Wscript.StdOut.WriteLine vbCrLf & "========================================================================="
Wscript.StdOut.WriteLine(" Script " & strTitle & " " & Now())
Wscript.StdOut.WriteLine "=========================================================================" & vbCrLf

objLogFile.WriteLine(vbCrLf & "=========================================================================")
objLogFile.WriteLine(" Script " & strTitle & " " & Now())
objLogFile.WriteLine("=========================================================================" & vbCrLf)

intUserRoleCount = 0
intMalformedAddresses = 0
intWellformedAddresses = 0
intCorrectedAddresses = 0


' Connect to AD RootDSE to get Partition Info...
Set objRootDSE = GetObject("LDAP://RootDSE")

If Err.Number = 0 Then
	
	' Make sure this is only run in XXX
	If InStr(LCase(objRootDSE.Get("defaultNamingContext")), "dc=XXX,") Then
		objLogFile.WriteLine("ERROR this script must only be run in the XXX Domain." & vbCrLf)
		WScript.Echo("ERROR this script must only be run in the XXX Domain.")
		objLogFile.Close
		WScript.Quit(1)
	End If
	
	' Create Search Parameters
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"

	Set objCommand.ActiveConnection = objConnection

	objCommand.CommandText = "<EDMS://OU=Users,OU=Accounts," & objRootDSE.Get("defaultNamingContext") & ">;(objectCategory=User);Name,distinguishedName;subtree"

	' Search for all Users
	Set objRecordSet = objCommand.Execute

		
		Do Until objRecordSet.EOF
	
			strUserDN = objRecordSet.Fields("distinguishedName").value
			strUserName = objRecordSet.Fields("Name").value
			
			Set objUser = GetObject("EDMS://" & strUserDN)
			intUserRoleCount = intUserRoleCount + 1
			strOldX400Addr = objUser.Get("textencodedoraddress")

			If inStr(strOldX400Addr, "p=mod1") Then
				intMalformedAddresses = intMalformedAddresses +1
				strNewX400Addr = Replace(strOldX400Addr, "p=XXX", "p=XXX")
				objUser.PutEx ADS_PROPERTY_DELETE, "textEncodedORAddress", strOldX400Addr
				objUser.SetInfo
				objUser.GetInfo
				
				objUser.PutEx ADS_PROPERTY_APPEND, "textEncodedORAddress", strNewX400Addr
				objUser.SetInfo
				
				If Err.Number = 0 Then
					objLogFile.WriteLine("Account Name: " & strUserName & " X400 address changed to " & strNewX400Addr)
				Else
					objLogFile.WriteLine("Could not ammend user: " & strUserName)
				End If
			Else
				intWellformedAddresses = intWellformedAddresses + 1
				objLogFile.WriteLine("Account Name: " & strUserDN & " fine, no changes needed.")
			End If
			objRecordset.movenext
		Loop

Else
	objLogFile.WriteLine("ERROR AD RootDSE Could Not be Contacted")
	Wscript.StdOut.WriteLine("ERROR AD RootDSE Could Not be Contacted")
End If

objLogFile.WriteLine("Command Completed.  " & vbcrlf & _
      intUserRoleCount & " Users/Roles processed. " & vbcrlf & _
      intWellformedAddresses & " addresses OK. " & vbcrlf  & _
      intCorrectedAddresses & " addresses corrected." & vbcrlf  & _
      intMalformedAddresses - intCorrectedAddresses & " address errors remaining.")

objLogFile.Close

' Display Output so User will Know Script Completed
WScript.Echo "Command Completed." & vbCRLF & _
	intUserRoleCount & " Users/Roles processed," & vbCRLF & _
    intWellformedAddresses & " addresses OK," & vbCRLF & _
    intCorrectedAddresses & " addresses corrected," & vbCRLF & _
    intMalformedAddresses - intCorrectedAddresses & " address errors remaining." & vbCRLF & vbCRLF & _
    "See Log File: " & strLogFilePath & " for further details."
	
WScript.quit
