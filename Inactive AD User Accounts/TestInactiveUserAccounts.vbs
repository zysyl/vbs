Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2

DisableUserAccount( "1ECB0F75664C85448CAA144FEB8D52B8" )
GenerateEmail( "1" )

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
Sub GenerateEmail( ByRef jobStats )

	Dim objEmail : Set objEmail = New EmailHelper
	Dim emailBody 
	
	Const EML_FROM = "unusedaccs@sportingbet.com"
	Const EML_TO = "stephen.hackett@sportingbet.com, hackett_s@hotmail.com"
	Const EML_SUBJECT = "Inactive User Accounts"
	emailBody = "" & VbCrLf & VbCrLf & _
				"Please note: this e-mail was sent from a notification-only address " & VbCrLf  & _
				"that cannot accept incoming e-mail. Please do not reply to this message."
	
	Call objEmail.SendMail( EML_FROM, EML_TO, EML_SUBJECT, emailBody )
	
	
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
	' @param    emlFrom		The originator of the email
	' @param    emlTo		The recipient of the email
	' @param    emlSubject  The subject line of the message
	' @param    emlBody		The actual email message
	'---------------------------------------------------------------------
	Public Sub SendMail(emlFrom, emlTo, emlSubject, emlBody)											
		m_objEmail.From = emlFrom
		m_objEmail.To = emlTo
		m_objEmail.ReplyTo = emlFrom
		m_objEmail.Subject = emlSubject
		m_objEmail.Textbody = emlBody
		m_objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		m_objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.43.4.8" 
		m_objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		m_objEmail.Configuration.Fields.Update
		m_objEmail.Send
	End Sub

End Class