Dim glb_ObjADConnection, glb_ObjADCommand, glb_objADRecordset

Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2
Const SQL_STATEMENT = "SELECT objectGuid, distinguishedName, displayName, department, samAccountName, lastLogonTimestamp, description, userAccountControl FROM 'LDAP://DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='user' AND objectClass='user' ORDER BY samAccountName"					  

' Initialise ADO 
Call InitialiseADO

' Query AD
Set glb_objADRecordset = glb_ObjADCommand.Execute

' Test recordset is not empty
If glb_objADRecordset.EOF Or glb_objADRecordset.BOF Then
	WScript.Echo "No results returned"
	WScript.Quit(-1)
End If

Do Until glb_objADRecordset.EOF
	Dim accountEnabled : accountEnabled = IsAccountEnabled( glb_objADRecordset.Fields("distinguishedName"), glb_objADRecordset.Fields("userAccountControl") )
	Dim objLastLogon   : objLastLogon = glb_objADRecordset.Fields("lastLogonTimestamp")
	Dim lastLogon	   : lastLogon = ConvertADLastLogonToDate(objLastLogon)	
	Dim department : department = glb_objADRecordset.Fields("department")
	Dim description  : description = glb_objADRecordset.Fields("description")
	Dim displayName : displayName = glb_objADRecordset.Fields("displayName")
	Dim samAccountName : samAccountName = glb_objADRecordset.Fields("samAccountName")
	
	' Only want Enabled Accounts & not Generic
	If IsAccountGeneric( glb_objADRecordset.Fields("distinguishedName") ) = False Then
	
		If accountEnabled = True Then	
			WScript.Echo samAccountName & vbTab & displayName & vbTab & department & vbTab & lastLogon		
		End If
	
	End If
		
	glb_objADRecordset.MoveNext
	
Loop

Function IsAccountGeneric( distinguishedName )

	If Instr( distinguishedName, "Shared Mailboxes") Then
		IsAccountGeneric = True	
	ElseIf InStr(distinguishedName, "Shared Accounts" ) Then
		IsAccountGeneric = True
	ElseIf InStr(distinguishedName, "Service Accounts" ) Then
		IsAccountGeneric = True
	ElseIf InStr( distinguishedName, "Application Accounts" ) Then
		IsAccountGeneric = True
	ElseIf InStr( distinguishedName, "Microsoft Exchange System Objects" ) Then
		IsAccountGeneric = True
	ElseIf InStr( distinguishedName, "Unity" ) Then
		IsAccountGeneric = True
	End If

End Function

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

Function IsAccountEnabled( dn, userAccountControl )
	    
    If userAccountControl = "512" Or _
    	userAccountControl = "544" Or _
    	 userAccountControl = "66048" Or _
    	   userAccountControl = "66080" Then    	  
    	       	      
    	IsAccountEnabled = True
    
    ElseIf userAccountControl = "514" Or _
    		userAccountControl = "546" Or _
    		 userAccountControl = "66050" Or _
    		 userAccountControl = "66082" Then
    		 
    	
    	IsAccountEnabled = False
    
    Else
    	' This catches unknown userAccountControl values
    	IsAccountEnabled = False
    	'Call WriteToLog("UnknownUACDetected" & FormatLogDate & ".log", "Unknown UserAccountControl Value Detected: " & _
    	'														userAccountControl & vbTab & dn )    	    
    End If
    
    
End Function

Function ConvertADLastLogonToDate( objLastLogon )

	Dim lastLogon
	
	If Not IsNull (objLastLogon) Then
	
		lastLogon = objLastLogon.HighPart * (2^32) + objLastLogon.LowPart
		lastLogon = lastLogon / (60 * 10000000)
		lastLogon = lastLogon / 1440
		lastLogon = lastLogon + #1/1/1601#
	
		ConvertADLastLogonToDate = lastLogon	
	Else
		
		ConvertADLastLogonToDate = "EMPTY"
	
	End If
	
	
End Function