Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2

Dim glb_ObjADConnection
Dim glb_ObjADCommand
Dim glb_objADRecordset

Const SQL_STATEMENT = "SELECT distinguishedName, samAccountName, lastLogonTimestamp, userAccountControl, displayName, mail, description, department, telephoneNumber, ipPhone, mobile FROM 'LDAP://DC=Sbet-EMEA,DC=ADS' WHERE objectClass='user' AND objectCategory='user' "

' Initialise ADO 
Call InitialiseADO
glb_totalActionable = 0

' Query AD
Set glb_objADRecordset = glb_ObjADCommand.Execute

' Test recordset is not empty
If glb_objADRecordset.EOF Or glb_objADRecordset.BOF Then
	WScript.Echo "No results returned"
	WScript.Quit(-1)
End If

' ### Main ###
Do Until glb_objADRecordset.EOF
	' Flags indicating account status
	Dim accountEnabled 
	accountEnabled = IsAccountEnabled( glb_objADRecordset.Fields("distinguishedName"), glb_objADRecordset.Fields("userAccountControl") )
	
	Dim accountGeneric 
	accountGeneric = IsAccountGeneric( glb_objADRecordset.Fields("distinguishedName") )
	
	Dim objLastLogon   
	objLastLogon = glb_objADRecordset.Fields("lastLogonTimestamp")
		
	Dim lastLogon	   
	lastLogon = ConvertADLastLogonToDate(objLastLogon)	
	
		
	' Ignore generic accounts
	If accountGeneric Then									
		'WScript.Echo "Ignore " & glb_objADRecordset.Fields("samAccountName")
	Else
		' Only action enabled accounts
		If accountEnabled = True Then	
			Dim objDescription
			objDescription = glb_objADRecordset.Fields("description").Value
			
			Dim description
						
			If IsArray(objDescription) = True Then				
				description = Join( glb_objADRecordset.Fields("description") )
			Else							
				description = ""
			End If
					
			WScript.Echo glb_objADRecordset.Fields("samAccountName") & vbTab &_
							glb_objADRecordset.Fields("mail")  & vbTab &_
							 glb_objADRecordset.Fields("displayName") & vbTab &_		
							  glb_objADRecordset.Fields("department") & vbTab &_		
							   glb_objADRecordset.Fields("telephoneNumber") & vbTab &_		
							    glb_objADRecordset.Fields("ipPhone") & vbTab &_		
							   	 glb_objADRecordset.Fields("mobile") & vbTab &_							   	  
							   	  description
							   							 
			'Call RenameDepartment( BytesToHexString( glb_objADRecordset.Fields("objectGUID") ) )
		End If		
	End If	

	' Move the recordset cursor forward one record
	glb_objADRecordset.MoveNext

Loop

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
' NAME: IsGenericAccount
' PURPOSE:  Determine if an account is a generic account based On
'			distinguishedName
' @param    distinguishedName	The raw AD lastlogontimestamp value
'---------------------------------------------------------------------
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

'---------------------------------------------------------------------
' NAME: IsAccountDisabled
' PURPOSE:  Determine the current status of an AD account based on 
'			useraccountControl attribute
' @param    userAccountControl	The raw AD userAccountControl value
' NOTE:		Known userAccountControl Values checked:
'			+ 512		Enabled Account
' 			+ 514		Disabled Account
'			+ 544		Enabled, Password Not Required
'			+ 546		Disabled, Password Not Required
'			+ 66048		Enabled, Password Doesn't Expire
'			+ 66050		Disabled, Password Doesn't Expire
'			+ 66080		Enabled, Password Doesn't Expire & Not Required
'			+ 66082		Disabled, Password Doesn't Expire & Not Required
'
'			Known userAccountControl Values NOT checked:
'			+ 262656	Enabled, Smartcard Required
'			+ 262658	Disabled, Smartcard Required
'			+ 262688	Enabled, Smartcard Required, Password Not Required
'			+ 262690	Disabled, Smartcard Required, Password Not Required
'			+ 328192	Enabled, Smartcard Required, Password Doesn't Expire
'			+ 328194	Disabled, Smartcard Required, Password Doesn't Expire
'			+ 328224	Enabled, Smartcard Required, Password Doesn't Expire & Not Required
'			+ 328226	Disabled, Smartcard Required, Password Doesn't Expire & Not Required
'			Unknown values logged to UnknownUACDetected.log
'			http://support.microsoft.com/kb/305144
'			http://www.netvision.com/ad_useraccountcontrol.php
'---------------------------------------------------------------------
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
    	WScript.Echo "Unknown UserAccountControl Value Detected: " & _
    															userAccountControl & vbTab & dn 
    End If
    
    
End Function


'---------------------------------------------------------------------
' NAME: ConvertADLastLogonToDate
' PURPOSE:  Convert a large integer value to a datestamp 
' @param    objLastLogon	The raw AD lastlogontimestamp value
'---------------------------------------------------------------------
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