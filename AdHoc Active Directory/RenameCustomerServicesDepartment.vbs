Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2

Dim glb_ObjADConnection
Dim glb_ObjADCommand
Dim glb_objADRecordset
Dim glb_totalActionable

Const SQL_STATEMENT = "SELECT objectGuid, distinguishedName, samAccountName, lastLogonTimestamp, userAccountControl, displayName, mail, description, department FROM 'LDAP://DC=Sbet-EMEA,DC=ADS' WHERE objectClass='user' AND description='Customer Services*' OR department='Customer Services*' ORDER BY sn"

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
		WScript.Echo "Ignore " & glb_objADRecordset.Fields("samAccountName")
	Else
		' Only action enabled accounts
		If accountEnabled = True Then
			glb_totalActionable = glb_totalActionable + 1
			WScript.Echo "Process " & glb_objADRecordset.Fields("samAccountName") 
			Call RenameDepartment( BytesToHexString( glb_objADRecordset.Fields("objectGUID") ) )
		End If		
	End If	

	' Move the recordset cursor forward one record
	glb_objADRecordset.MoveNext

Loop

WScript.Echo VbCrLf & VbCrLf & "SUMMARY: Total Actionable accounts: " & glb_totalActionable 

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
' NAME: IsAccountInactive
' PURPOSE:  Determine if an Active Directory account last logged
'			in within the last 90 days
' @param    objLastLogon	The raw AD lastlogontimestamp value
'---------------------------------------------------------------------
Function IsAccountInactive( objLastLogon )

	Dim lastLogonDateStamp
	Dim lastLogonInDays
	
	lastLogonDateStamp = ConvertADLastLogonToDate( objLastLogon )	
	
	If lastLogonDateStamp <> "EMPTY" Then	
	
		lastLogonInDays = DateDiff("d",lastLogonDateStamp, Now)
		
		If lastLogonInDays >= 90 Then
			IsAccountInactive = True
		Else
			IsAccountInactive = False
		End If
	
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


'---------------------------------------------------------------------
' NAME: DisableUserAccount
' PURPOSE:  Disable a user account in AD and write to the
'			objects description field
' @param    guid		The unique id of the account object
'---------------------------------------------------------------------
Function RenameDepartment( guid )
	
	'F8F84422061EBF4F90E46AFDB6CA34F3 	 		
	
	' Bind to user account
	Dim objUserAccount : Set objUserAccount = GetObject("LDAP://<GUID=" & guid & ">")				
		
	WScript.Echo "  Old: " & objUserAccount.Department
	
	' Clear out old description field entry
	objUserAccount.PutEx ADS_PROPERTY_CLEAR, "department", 0	
	
	' Write date disabled in description
	objUserAccount.Put "department", "Customer Services"
	
	' Save Changes
	objUserAccount.SetInfo	
	
	WScript.Echo "  New: " & objUserAccount.Department

End Function