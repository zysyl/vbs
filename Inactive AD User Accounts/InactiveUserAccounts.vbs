'=====================================================================
' NAME:		InactiveUserAccounts.vbs
'
' PURPOSE:	Disable Active Directory User accounts that have not
'			logged on within the last 90 days.
'
' USEAGE: 	CSCRIPT InactiveUserAccounts.vbs, needs to be run with
'			raised privileged account.
'
' NOTE:
'
' MODIFIED: (1) 16-Sep-2011 Version 1.0.0: Initial Release
'			(2) 19-Sep-2011 Removed Check for null samAccountName
'			(3) 19-Sep-2011 Added logic to handle null lastLogonTimestamp
'			(4) 20-Sep-2011 Improved check for matching criteria to disable
'	   
'=====================================================================	
Option Explicit

Dim glb_ObjADConnection, glb_ObjADCommand, glb_objADRecordset
Dim glb_Logfilename : glb_Logfilename = Replace(WScript.ScriptName, ".vbs", "-" & FormatLogDate & ".csv" )
Dim glb_Errfilename : glb_Errfilename = Replace(WScript.ScriptName, ".vbs", "-" & FormatLogDate & ".err" )
Dim glb_totalAccounts, glb_totalGeneric, glb_totalEnabled, glb_totalDisabled, glb_totalToDisable
Dim glb_totalEnabledAndGeneric, glb_NotGenericNeverLoggedIn, glb_totalValid
Dim jobStats, jobSuccess

jobSuccess = True

Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2
Const SQL_STATEMENT = "SELECT objectGuid, distinguishedName, samAccountName, lastLogonTimestamp, userAccountControl FROM 'LDAP://DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='user' AND objectClass='user' ORDER BY samAccountName"					  

' Initialise Counters
Call InitialiseCounters

' Initialise ADO 
Call InitialiseADO

' Query AD
Set glb_objADRecordset = glb_ObjADCommand.Execute

' Test recordset is not empty
If glb_objADRecordset.EOF Or glb_objADRecordset.BOF Then
	WScript.Echo "No results returned"
	WScript.Quit(-1)
End If

glb_totalAccounts = glb_objADRecordset.RecordCount

' Write Headers to the logfile
Call WriteToLog (glb_Logfilename, "Account GUID" &vbTab& "Distinguished Name" &vbTab& "Username" &vbTab& "Account Enabled" &vbTab& "Remark" &vbTab& "Last Logon")

' ### 00. Script Body
Do Until glb_objADRecordset.EOF
		
	Dim accountEnabled : accountEnabled = IsAccountEnabled( glb_objADRecordset.Fields("distinguishedName"), glb_objADRecordset.Fields("userAccountControl") )
	Dim accountGeneric : accountGeneric = IsAccountGeneric( glb_objADRecordset.Fields("distinguishedName") )
	Dim objLastLogon   : objLastLogon = glb_objADRecordset.Fields("lastLogonTimestamp")
	Dim lastLogon	   : lastLogon = ConvertADLastLogonToDate(objLastLogon)	
	
	' ### 01. Update counters
	If accountEnabled Then
		glb_totalEnabled = glb_totalEnabled + 1
		
		If accountGeneric Then
			glb_totalEnabledAndGeneric = glb_totalEnabledAndGeneric + 1					
		Else
			If lastLogon = "EMPTY" Then
				glb_NotGenericNeverLoggedIn = glb_NotGenericNeverLoggedIn + 1
			Else
				glb_totalValid = glb_totalValid + 1
			End If					
		End If
		
	Else
		glb_totalDisabled = glb_totalDisabled + 1					
	End If
	
	
	' ### 02. Ignore Generic Accounts			
	If accountGeneric Then
		glb_totalGeneric = glb_totalGeneric + 1												  							
	Else
		
		' ### 03. Criteria to disable						
		' Check Account meets inactive criteria (not logged on for 90 days) and is currently enabled
		If IsAccountInactive( objLastLogon ) And accountEnabled = True Then	
		
			On Error Resume Next
				
				Call DisableUserAccount( BytesToHexString( glb_objADRecordset.Fields("objectGUID") ) )
				If Err.Number = 0 Then					
					glb_totalToDisable = glb_totalToDisable + 1
					Call WriteToLog (glb_Logfilename, BytesToHexString( glb_objADRecordset.Fields("objectGUID") ) & vbTab & _
									   glb_objADRecordset.Fields("distinguishedName") & vbTab & _
									    glb_objADRecordset.Fields("samAccountName") & vbTab & _
									     accountEnabled & vbTab & "Account has been disabled" & vbTab & _
					  					  lastLogon )
				Else
					Call WriteToLog	( glb_Errfilename, "[ERROR] :: Trying to disable account " & vbTab & _
									   BytesToHexString( glb_objADRecordset.Fields("objectGUID") ) & vbTab & _
								   		glb_objADRecordset.Fields("distinguishedName") & vbTab & _
								    	 glb_objADRecordset.Fields("samAccountName") )	
								    	 
					jobSuccess = False
												
				End If	
				
			Err.Clear			
																																					
		End If			
	
	
	End If
	
	' Move the recordset cursor forward one record
	glb_objADRecordset.MoveNext
	
Loop

' Build up statistics for this job to send in the email
If jobSuccess = True Then
	jobStats = Date & ": Inactive Active Directory Accounts : " & glb_totalToDisable & VbCrLf & VbCrLf
	jobStats = jobStats & "For a detailed list of all accounts that have been disabled please see " & glb_Logfilename & " attached." & VbCrLf & VbCrLf 	
	Call GenerateEmail( jobStats, "W:\Scheduled Tasks\Inactive AD Users\" & glb_Logfilename )
Else
	jobStats = Date & ": Inactive Active Directory Accounts : " & glb_totalToDisable & VbCrLf & VbCrLf
	Call GenerateErrorEmail( jobStats, "W:\Scheduled Tasks\Inactive AD Users\" & glb_Errfilename )
End If

'=========================== END OF SCRIPT =================================




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
' NAME: InitialiseCounters
' PURPOSE:  Utility method to set counters to zero
' @param    
'---------------------------------------------------------------------
Sub InitialiseCounters

	glb_totalAccounts = 0
	glb_totalGeneric = 0
	glb_totalEnabled = 0
	glb_totalDisabled = 0
	glb_totalToDisable = 0
	glb_totalEnabledAndGeneric = 0
	glb_NotGenericNeverLoggedIn = 0
	glb_totalValid = 0

End Sub



'---------------------------------------------------------------------
' NAME: WriteToLog
' PURPOSE:  Create a text file to use for logging script output/events
'			and write the log entry
' @param    logFilename		The log file to write to
' @param    messageText		The message to log
'---------------------------------------------------------------------
Sub WriteToLog( logFilename, messageText )
	
	Dim localLogFile, objFSO
	
	Set objFSO = CreateObject("Scripting.Filesystemobject")
	If objFSO.FileExists(logFilename) Then
		Set localLogFile = objFSO.OpenTextFile(logFilename, 8, True)
	Else		
		Set localLogFile = objFSO.CreateTextFile(logFilename, True)
	End If
	localLogFile.WriteLine messageText
	localLogFile.Close

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
	ElseIf Instr( distinguishedName, "Generic Accounts") Then
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
    	Call WriteToLog("UnknownUACDetected" & FormatLogDate & ".log", "Unknown UserAccountControl Value Detected: " & _
    															userAccountControl & vbTab & dn )    	    
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
' NAME: FormatLogDate
' PURPOSE:   Returns a string representing the execution date 
'			 of the script in the form YYYY-MM-DD
' @param    
'---------------------------------------------------------------------
Private Function FormatLogDate()
	
	Dim nYear, nMonth, nDate, nDay
	
	nYear = Year(Now)
	nMonth = Month(Now)
	nDay = Day(Now)
	
	If Len(nMonth) = 1 Then
		nMonth = "0" & nMonth
	End If
	
	If Len(nDay) = 1 Then
		nDay = "0" & nDay	
	End If
	
	FormatLogDate = nYear & nMonth & nDay	
	
End Function



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
' NAME: DisableUserAccount
' PURPOSE:  Disable a user account in AD and write to the
'			objects description field
' @param    guid		The unique id of the account object
'---------------------------------------------------------------------
Function DisableUserAccount( guid )
 	 		
	' Bind to user account
	Dim objUserAccount : Set objUserAccount = GetObject("LDAP://<GUID=" & guid & ">")	
			
	' Disable it
	objUserAccount.Put "userAccountControl", ADS_UF_ACCOUNTDISABLE
		
	' Clear out old description field entry
	objUserAccount.PutEx ADS_PROPERTY_CLEAR, "description", 0
	
	' Write date disabled in description
	objUserAccount.Put "description", "### ACCOUNT DISABLED - " & Date & " - ###"
	
	' Save Changes
	objUserAccount.SetInfo

End Function




'---------------------------------------------------------------------
' NAME: GenerateEmail
' PURPOSE:  Wrapper Routine to build an email message
' @param    
'---------------------------------------------------------------------
Sub GenerateEmail( ByRef jobStats, ByVal attachment )

	Dim objEmail : Set objEmail = New EmailHelper
	Dim emailBody 
	
	Const EML_FROM = "inactiveaccs@sportingbet.com"
	Const EML_TO = "moorfields.security@sportingbet.com, HelpdeskGlobal@sportingbet.com"
	'Const EML_TO = "stephen.hackett@sportingbet.com"
	Const EML_SUBJECT = "Inactive User Accounts"
	emailBody = jobStats & VbCrLf & VbCrLf & _
				"Please note: this e-mail was sent from a notification-only address " & VbCrLf  & _
				"that cannot accept incoming e-mail. Please do not reply to this message."
	
	Call objEmail.SendMail( EML_FROM, EML_TO, EML_SUBJECT, emailBody, attachment )
	
	
End Sub




'---------------------------------------------------------------------
' NAME: GenerateErrorEmail
' PURPOSE:  Wrapper Routine to build an email message on job error
' @param    
'---------------------------------------------------------------------
Sub GenerateErrorEmail( ByRef jobStats, ByVal attachment )

	Dim objEmail : Set objEmail = New EmailHelper
	Dim emailBody 
	
	Const EML_FROM = "inactiveaccs@sportingbet.com"	
	Const EML_TO = "stephen.hackett@sportingbet.com"
	Const EML_SUBJECT = "Inactive User Accounts: FAILED"
	emailBody = jobStats & VbCrLf & VbCrLf & _
				"Please note: this e-mail was sent from a notification-only address " & VbCrLf  & _
				"that cannot accept incoming e-mail. Please do not reply to this message."
	
	Call objEmail.SendMail( EML_FROM, EML_TO, EML_SUBJECT, emailBody, attachment )
	
	
End Sub



'---------------------------------------------------------------------
' NAME: EmailHelper
' PURPOSE:  Utility class to generate email
' @param    
'---------------------------------------------------------------------
Class EmailHelper

	Private m_objShell
	Private m_objEmail
	
	'---------------------------------------------------------------------
	' NAME: Default Constructor
	' PURPOSE:  Initialise class
	' @param    
	'---------------------------------------------------------------------
	Private Sub class_initialize
		Set m_objShell = CreateObject("WScript.Shell")
		Set m_objEmail = CreateObject("CDO.Message")
	End Sub	
	
	'---------------------------------------------------------------------
	' NAME: SendMail
	' PURPOSE:  Actually send an email message
	' @param    emlFrom			The originator of the email
	' @param    emlTo			The recipient of the email
	' @param    emlSubject  	The subject line of the message
	' @param    emlBody			The actual email message
	' @param    emalAttachment	An email attachment 
	'---------------------------------------------------------------------
	Public Sub SendMail(emlFrom, emlTo, emlSubject, emlBody, emlAttachment)											
		m_objEmail.From = emlFrom
		m_objEmail.To = emlTo
		m_objEmail.ReplyTo = emlFrom
		m_objEmail.Subject = emlSubject
		m_objEmail.Textbody = emlBody
		m_objEmail.AddAttachment emlAttachment					
		m_objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		m_objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.43.4.8" 
		m_objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		m_objEmail.Configuration.Fields.Update
		m_objEmail.Send
	End Sub

End Class