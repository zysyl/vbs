'Option Explicit

Dim arrUsers			' contains the input from text file
Dim glb_ObjADConnection
Dim glb_ObjADCommand 
Dim glb_objADRecordset


Const INPUT_FILE_PATH = ".\usernames.txt"		' the path to the file
Const FOR_READING = 1
Const FOR_WRITING = 2

Const ADS_SCOPE_SUBTREE = 2	
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2
'Const SQL_STATEMENT = "SELECT samAccountName, distinguishedName FROM 'LDAP://OU=Users,OU=LON,OU=UK,DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='user' ORDER BY samAccountName"					  
'Const SQL_STATEMENT = "SELECT samAccountName, distinguishedName FROM 'LDAP://OU=Finsoft,OU=External,OU=Generic Accounts,OU=TWH,OU=UK,DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='user' ORDER BY samAccountName"					  
'Const SQL_STATEMENT = "SELECT samAccountName, distinguishedName FROM 'LDAP://OU=GAVS,OU=External,OU=Generic Accounts,OU=TWH,OU=UK,DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='user' ORDER BY samAccountName"					  
Const SQL_STATEMENT = "SELECT name, dnsHostname, operatingSystem, distinguishedName FROM 'LDAP://OU=Terminated,OU=Old Computer Accounts,OU=Domain Management,DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='computer' ORDER BY cn"					  

Call ReadUserAndExtensionFromFile()

' ==== END OF SCRIPT ===


Sub QueryADViaSql()		

	' Initialise ADO 
	Call InitialiseADO
	
	' Query AD
	Set glb_objADRecordset = glb_ObjADCommand.Execute
	
	' Test recordset is not Empty
	If glb_objADRecordset.EOF Or glb_objADRecordset.BOF Then
		WScript.Echo "No results returned"
		WScript.Quit(-1)
	End If
	
	Do Until glb_objADRecordset.EOF
	
		'Call PrintDetails( glb_objADRecordset.Fields("distinguishedName") )
		Dim hostname
		Dim dnsHostname
		Dim ipAddress
		Dim os
		Dim rverseLookup
		
		' ### Hostname
		If IsNull( glb_objADRecordset.fields("name") ) = True Then
			hostname = "EMPTY"
		Else
			hostname = glb_objADRecordset.fields("name")
		End If
		
		
		' ### DNS Hostname
		If IsNull( glb_objADRecordset.fields("dNSHostname") ) = True  Then
			dnsHostname = "EMPTY"
		Else
			dnsHostname = glb_objADRecordset.fields("dNSHostname")
		End If
		
		
		' ### IP Address Lookup
		If dnsHostname = "EMPTY" Then
			ipAddress = "EMPTY"			
		Else
			Dim online
			ipAddress = NSlookup( dnsHostname )		
			If Not ipAddress = "EMPTY" Then
				online = PingIP( ipAddress )
				reverseLookup = NSlookup( ipAddress )
			Else
				online = "EMPTY"
				reverseLookup = "EMPTY"				
			End If					
		End If
		
		
		' ### OS
		If IsNull( glb_objADRecordset.fields("operatingSystem") ) = True Then
			os = "EMPTY"
		Else
			os = glb_objADRecordset.fields("operatingSystem")
		End If
				
		' Spit it out
		WScript.Echo hostname & vbTab & dnsHostname & vbTab & ipAddress & vbTab & online & vbTab & reverseLookup & vbTab & os
		
		
		
		glb_objADRecordset.MoveNext											
		
	Loop
	
End Sub


Sub DisableUsersInAGroup

	Dim objAdLookup : Set objAdLookup = New ADLookup	
	Dim objGroup : objGroup = "CN=VPN UK Sportingbet,OU=VpnAccess,OU=Groups,OU=TWH,OU=UK,DC=Sbet-EMEA,DC=ADS"
	Dim arrMembers : arrMembers = objAdLookup.GetGroupMembers( objGroup )
	Dim arrUser
	Dim x
	x =0
	For i=LBound(arrMembers) To UBound(arrMembers)		
		Call DisableUserAccount( arrMembers(i) )			
	Next		
	WScript.Echo x
End Sub


Sub ReadUserAndExtensionFromFile()

	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Try to read in the content of INPUT_FILE_PATH
	If objFSO.FileExists(INPUT_FILE_PATH) = false Then
		WScript.Echo "Trying to read file " & INPUT_FILE_PATH & VbCrLf & "File not found"
		WScript.quit
	Else
		Set objInFile = objFSO.OpenTextFile(INPUT_FILE_PATH, FOR_READING)
		arrUsers = Split(objInFile.ReadAll, VbCrLf)
		objInFile.Close
	End If	
	
	' Loop the file and query AD for the users GUID
	For i=LBound(arrUsers) To UBound(arrUsers)	
			On Error Resume Next
				'WScript.Echo arrUsers(i)
				Dim username : username = split(arrUsers(i),",")(0)
				'Dim ipPhone : ipPhone = split(arrUsers(i),",")(1)
				Dim objGuid : objGuid = GetGuidFor(username)		
				'Wscript.Echo  objGuid
				Call UpdateIpPhone( objGuid, ipPhone  )		
			Err.Clear	
	Next

End Sub


Sub DisableUsrAccountsViaSqlQuery()		
	' Initialise ADO 
	Call InitialiseADO
	
	' Query AD
	Set glb_objADRecordset = glb_ObjADCommand.Execute
	
	' Test recordset is not Empty
	If glb_objADRecordset.EOF Or glb_objADRecordset.BOF Then
		WScript.Echo "No results returned"
		WScript.Quit(-1)
	End If
	
	Do Until glb_objADRecordset.EOF
		Call DisableUserAccount( glb_objADRecordset.Fields("distinguishedName") )
	
		glb_objADRecordset.MoveNext											
		
	Loop
	
End Sub



' ===== End Of Script ===

'
'
Function PingIP ( ByVal ip )

Dim objReg
Dim objPingStatusRet
Dim objPing32

	' Setup wmi object
	Set objWMI = GetObject("winmgmts:\\.\root\cimv2")

	' Check the machine is online
	Set objPing32 = objWMI.ExecQuery("SELECT StatusCode, ProtocolAddress FROM Win32_PingStatus " & _
								 "WHERE Address = '" & ip & "'")
	For Each objPingStatusRet In objPing32
		If IsNull(objPingStatusRet.StatusCode) Or objPingStatusRet.StatusCode <> 0 Then
			PingIP = "Offline"
		Else
			PingIP = "Online"
		End If						
	Next

End Function



' Script source:
' https://groups.google.com/group/microsoft.public.scripting.vbscript/msg/a465907f8dc6e265
Function NSlookup( sHost ) 
    ' Both IP address and DNS name is allowed 
    ' Function will return the opposite 
    Set oRE = New RegExp 
    oRE.Pattern = "^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$" 
    bInpIP = False 
    If oRE.Test(sHost) Then 
        bInpIP = True 
    End If 
    Set oShell = CreateObject("Wscript.Shell") 
    Set oFS = CreateObject("Scripting.FileSystemObject") 
    sTemp = oShell.ExpandEnvironmentStrings("%TEMP%") 
    sTempFile = sTemp & "\" & oFS.GetTempName     
    
    'Run NSLookup via Command Prompt 
    'Dump results into a temp text file 
    oShell.Run "%ComSpec% /c nslookup.exe " & sHost & " >" & sTempFile, 0, True 
    
    'Open the temp Text File and Read out the Data 
    Set oTF = oFS.OpenTextFile(sTempFile) 
    
    'Parse the text file 
    Do While Not oTF.AtEndOfStream 
        sLine = Trim(oTF.Readline) 
        If LCase(Left(sLine, 5)) = "name:" Then 
            sData = Trim(Mid(sLine, 6)) 
            If Not bInpIP Then 
                'Next line will be IP address(es) 
                'Line can be prefixed with "Address:" or "Addresses": 
                aLine = Split(oTF.Readline, ":") 
                sData = Trim(aLine(1)) 
            End If 
            Exit Do 
        End If 
    Loop 
    
    'Close it 
    oTF.Close 
    
    'Delete It 
    oFS.DeleteFile sTempFile 
    
    If Lcase(TypeName(sData)) = LCase("Empty") Then 
        NSlookup = "EMPTY" 
    Else     	
        NSlookup = sData 
    End If 
    
End Function 


'---------------------------------------------------------------------
' NAME: InitialiseADO
' PURPOSE:  Utility method to setup ADO
' @param    
'---------------------------------------------------------------------
Sub InitialiseADO
	' Initialise ADO objects
	Set glb_ObjADConnection = CreateObject("ADODB.Connection")
	Set glb_ObjADCommand = CreateObject("ADODB.Command")
	Set glb_objADRecordset = CreateObject("ADODB.Recordset")
	
	' Configure AD Connection object
	With glb_ObjADConnection
		.Provider = "ADsDSOObject"
		.Open "Active Directory Provider"
	End With
	
	' Configure ADO command object
	Set glb_ObjADCommand.ActiveConnection = glb_ObjADConnection					  
	With glb_ObjADCommand
		.Properties("Page Size") = 1000
		.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
		.CommandText =  SQL_STATEMENT
	End With
End Sub



'---------------------------------------------------------------------
' NAME: DisableUserAccountsFromFile
' PURPOSE:  Read in a list of username from a file, get their guid
'			bind to the account disable it and update the description
' @param    
'---------------------------------------------------------------------
Sub DisableUserAccountsFromFile()

	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Try to read in the content of INPUT_FILE_PATH
	If objFSO.FileExists(INPUT_FILE_PATH) = false Then
		WScript.Echo "Trying to read file " & INPUT_FILE_PATH & VbCrLf & "File not found"
		WScript.quit
	Else
		Set objInFile = objFSO.OpenTextFile(INPUT_FILE_PATH, FOR_READING)
		arrUsers = Split(objInFile.ReadAll, VbCrLf)
		objInFile.Close
	End If	
	
	' Loop the file and query AD for the users GUID
	For i=LBound(arrUsers) To UBound(arrUsers)	
			'On Error Resume Next
				Dim objGuid : objGuid = GetGuidFor(arrUsers(i))		
				'Wscript.Echo BytesToHexString( objGuid )					
				Call DisableUserAccount( BytesToHexString( objGuid ) )		
			'Err.Clear	
	Next

End Sub



'---------------------------------------------------------------------
' NAME: BytesToHexString
' PURPOSE:  Utility script used to convert bytes array to Hexadecimal
'			string
' @param    
'---------------------------------------------------------------------
Function BytesToHexString(aBytes)
Dim i
BytesToHexString = ""
For i = 1 To LenB(aBytes)
	BytesToHexString = BytesToHexString & Right("0" & Hex(AscB(MidB(aBytes, i, 1))), 2)
Next
End Function



Function GetGuidFor(ByVal cn)
    ' Function:     GetGuidFor
    ' Description:  Returns GUID for user ID (CN)
    ' Parameters:   ByVal cn 
    ' Returns:      Returns GUID for user ID (CN)
   
    Dim oRootDSE, oConnection, oCommand, oRecordSet
        
    'On Error Resume Next
    
    Set oRootDSE = GetObject("LDAP://rootDSE")
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Open "Provider=ADsDSOObject;"
    Set oCommand = CreateObject("ADODB.Command")
    oCommand.ActiveConnection = oConnection
    oCommand.CommandText = "<LDAP://" & oRootDSE.get("defaultNamingContext") & _
    ">;(&(objectClass=User)(samAccountName=" & cn & "));distinguishedName;subtree"
    Set oRecordSet = oCommand.Execute       
    
    If oRecordSet.EOF Or oRecordSet.BOF Then
    	' the result set is empty
    	'GetGuidFor = cn & ",NOT FOUND"
    	WScript.Echo "ERROR " & vbTab & oCommand.CommandText
    	Err.Raise (999)
    Else
    	oRecordSet.MoveFirst
    	GetGuidFor = oRecordSet.Fields("distinguishedName")
    End If

  	'On Error GoTo 0 
    oConnection.Close
    
    'Set oRecordSet = Nothing
    'Set oCommand = Nothing
    'Set oConnection = Nothing
    'Set oRootDSE = Nothing
    
End Function

'---------------------------------------------------------------------
' NAME: DisableUserAccount
' PURPOSE:  Disable a user account in AD and write to the
'			objects description field
' @param    guid		The unique id of the account object
'---------------------------------------------------------------------
Function DisableUserAccount( dn )
 	
	' Bind to user account
	Dim objUserAccount : Set objUserAccount = GetObject("LDAP://" & dn)	
	
	'WScript.Echo objUserAccount.samAccountName & vbTab & objUserAccount.DisplayName & vbTab & objUserAccount.Description & vbTab & objUserAccount.Department
	WScript.Echo objUserAccount.samAccountName 
	
	' Disable it
	'objUserAccount.Put "userAccountControl", ADS_UF_ACCOUNTDISABLE
		
	' Clear out old description field entry
	'objUserAccount.PutEx ADS_PROPERTY_CLEAR, "description", 0
	
	' Write date disabled in description
	'objUserAccount.Put "description", "### ACCOUNT DISABLED ###"
	
	objUserAccount.Put "company", "GAVS"
	
	' Save Changes
	objUserAccount.SetInfo

End Function


Function UpdateIpPhone( ByVal dn, ByVal ipPhone )
 	
	' Bind to user account
	Dim objUserAccount : Set objUserAccount = GetObject("LDAP://" & dn)	
	
	WScript.Echo objUserAccount.DisplayName & vbTab & objUserAccount.ipPhone
	
	' Write date disabled in description
	'objUserAccount.Put "ipPhone", ipPhone
	
	' Save Changes
	'objUserAccount.SetInfo

End Function


'---------------------------------------------------------------------
' NAME: DisableUserAccount
' PURPOSE:  Disable a user account in AD and write to the
'			objects description field
' @param    guid		The unique id of the account object
'---------------------------------------------------------------------
Function PrintDetails( dn)
	Dim objComputer : Set objComputer = GetObject("LDAP://" & dn)	
	
	' Spit out to screen
	WScript.Echo objComputer.Name & vbTab & objComputer.dNSHostName 
	
	' Query DNS
	WScript.Echo "DNS: " & NSlookup( objComputer.dNSHostName )
	
	
	If objComputer.dNSHostname = "" Then
		WScript.Echo "DNSHostname EMPTY"
	End If
					
End Function


' ********
' Active Directory Lookup Class
' ********
Class ADLookup

	Private m_ADS_SCOPE_SUBTREE
	Private m_objConnection, m_objCommand, m_objRecordSet
	Private m_Logger
	
	' ********
	' Class Constructor
	' ********
	Private Sub class_initialize
		m_ADS_SCOPE_SUBTREE = 2
	'	Set m_Logger = New ScheduleLogger		
		
		' Setup ADO connection properties		
		Set m_objConnection = CreateObject("ADODB.Connection")
		m_objConnection.Provider = "ADsDSOObject"
		m_objConnection.Open "Active Directory Provider"
		
		' Setup ADO Command object
		Set m_objCommand = CreateObject("ADODB.Command")			
		Set m_objCommand.ActiveConnection = m_objConnection
		m_objCommand.Properties("Page Size") = 1000
		m_objCommand.Properties("Searchscope") = m_ADS_SCOPE_SUBTREE 
		
				
		' Setup ADO Recordset
		Set m_objRecordSet = CreateObject("ADODB.Recordset")
	End Sub
	
	
	' ********
	' Get all of the members in an AD group
	' ********	
	Public Function GetGroupMembers(ByVal grpDistinguishedName)
	
	Dim objConnection, objCommand, objRecordSet
        
    'On Error Resume Next        	
	'	Call m_Logger.WriteLogEntry(Now & "," & "[INFO]" & "," & grpDistinguishedName & "'" )
	    m_objCommand.CommandText = "SELECT member FROM 'LDAP://" & grpDistinguishedName & "'"
        
        Set m_objRecordSet = m_objCommand.Execute
        
        If m_objRecordSet.BOF Or m_objRecordSet.EOF Then
        	GetGroupMembers = Null
        Else
        	GetGroupMembers = m_objRecordSet.Fields("member").value
        End If
        
  	'On Error GoTo 0 
    'm_objConnection.Close       
	    
	End Function

End Class