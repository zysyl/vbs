'=====================================================================
' NAME:		CommomFunctions.vbs
' PURPOSE:	Contains common utility functions shared by scripts
'=====================================================================	

'----------------------------------------------
' NAME: Log
' PURPOSE:  Log the message into the log file
' @param    strMessage	Log data to be logged into the file
'----------------------------------------------
Function Log(strMessage)
	Dim oStream: Set oStream = oFSO.OpenTextFile(sLogFileName, ForAppending, True)
	If Not oStream is Nothing then
		oStream.WriteLine "[" & Now() & "]" & strMessage
		oStream.Close
	Else
		WScript.Echo(strMessage)
	End if
	Set oStream = Nothing
End Function

'----------------------------------------------
' INFO: ValidateMail
' PURPOSE: Check for valid email address
' @param address	email address
' @return			boolean
'----------------------------------------------
Function ValidateMail(address)
	Dim objRegExpr
	Set objRegExpr = New RegExp

	objRegExpr.pattern = "^([0-9a-zA-Z]+[-._+&])*[0-9a-zA-Z]+@([-0-9a-zA-Z]+[.])+[a-zA-Z]{2,6}$"
	objRegExpr.Global = true
	objRegExpr.IgnoreCase = true

	if(objRegExpr.Test(address)) then
		ValidateMail = true
	else
		ValidateMail = false
	end if
End Function

'----------------------------------------------
' NAME: SendMessage
' PURPOSE: Sends a message to the specified user
' @param FromID		                From whom the email is being sent
' @param ToID		                'to' list of email recipients
' @param CC			                'cc'  list of email recipients
' @param BCC		                'Bcc' list of email recipients
' @param subject	                Subject of the message to be sent
' @param body		                Body of the message
' @param txtImageLogoFilePath       
' @param txtImageLogoFileName       
' @param IsBodyHTML                 Specify whether the Body is in HTML format or Plain Text
'----------------------------------------------
Sub SendMessage(FromID, ToID, CC, BCC, Subject, sBody, txtImageLogoFilePath, txtImageLogoFileName, bIsBodyHTML)
   Dim iConf
   Dim objSendMail
   Dim sTextBody, sHTMLBody
   Const cdoAnonymous = 0 'Do not authenticate - Anonymous
   Const cdoBasic = 1	'basic (clear-text) authentication
   Const cdoNTLM = 2	'NTLM authentication

	Const cdoRefTypeId = 0
	Const cdoRefTypeLocation = 1

    Set objSendMail = CreateObject("CDO.Message")
    objSendMail.Subject = Subject
    objSendMail.From =  FromID
    objSendMail.To = ToID
    objSendMail.Cc = CC
    objSendMail.Bcc = BCC
    If bIsBodyHTML = True Then
        objSendMail.HTMLBody = sBody
    Else
        objSendMail.TextBody = sBody
    End If
	objSendMail.AutoGenerateTextBody = true

    If bIsBodyHTML = True Then
	    'gives the capability to turn on and off the image attachment within the email notification
	    Dim sRemoveImageLogoAttachment: sRemoveImageLogoAttachment = GetVariableValue("RemoveImageLogoAttachment")
	    If trim(LCase(sRemoveImageLogoAttachment)) = "false" Then
	        If(oFSO.FileExists(txtImageLogoFilePath)) then
		        Call objSendMail.AddRelatedBodyPart( txtImageLogoFilePath, txtImageLogoFileName, cdoRefTypeId )
	        End If
	    End If    
    End if


	'configure the CDO mailer to send to the server specs defined in the web.config
	Set iConf = objSendMail.Configuration
	With iConf.Fields
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")		 = 2 ' SendUsingPort
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserver")       = GetVariableValue("SMTPServer")
		.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")   = GetVariableValue("SMTPServerPort")
		.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = GetVariableValue("SMTPAuthenticate")
		if (GetVariableValue("SMTPAuthenticate") <> cdoAnonymous) then
			.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = GetVariableValue("SendUserName")
			.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = decode64(GetVariableValue("SendPassword"))
		end if
		.Update
	End With

	WScript.Echo("INFO: Sending Mail to: " & ToID)
	Call Log("INFO: Sending Mail to: " & ToID & ". [CommonFunctions.vbs, SendMessage()]")
	objSendMail.send
	
	If Err.Number <> 0 Then
		Call Log("ERROR: After sending message [" & Err.Number & "] " & Err.Description & ". [CommonFunctions.vbs, SendMessage()]")
	End If
	
	Set objSendMail = Nothing
End Sub


'--------------------------------------------------
' NAME: GetFileName
' PURPOSE: Retrieve the file name from the complete path
'--------------------------------------------------
Function GetFileName(sVirtualPath)
	Dim iPos, sFile
	If trim(sVirtualPath) <> "" then
		iPos = InStrRev(sVirtualPath, "/")
		sFile = Right(sVirtualPath, len(sVirtualPath) - iPos)
		GetFileName	 = sFile
	Else
		GetFileName = ""
	End if
End Function

'-----------------------------------------------------------------
' NAME: GetMailContentFromTemplate
' PURPOSE: Retrieves the content from specified template file
' INPUTS: 
'-----------------------------------------------------------------
Function GetMailContentFromTemplate(strTemplateName)
	
	Dim sMessage
	Dim oFSO: Set oFSO = WScript.CreateObject("Scripting.FileSystemObject")
    Dim sTemplateFilePath: sTemplateFilePath = GetConfigPath() & "Templates\" & strTemplateName
    
    'ensure the template file exists
	If oFSO.FileExists(sTemplateFilePath)Then
		Dim oFile: Set oFile =	oFSO.openTextFile(sTemplateFilePath, 1, False)
		
		If Not oFile is nothing Then
		    'retrieving template contents
			sMessage = oFile.ReadAll
			GetMailContentFromTemplate = sMessage
			oFile.close
			Set oFile = nothing
		Else 
		    'an error occurred reading the template
            Call Log("ERROR: Could not read the contents of the template. [CommonFunctions.vbs, GetMailContentFromTemplate()]")
            Call Log("HINT: Ensure the script has the appropriate NTFS read file permissions for the following template " & sTemplateFilePath & ". [CommonFunctions.vbs, GetMailContentFromTemplate()]")
		End If
	Else
	    'path to the template was incorrect or template doesn't exist
		Call Log("ERROR: Could not retrieve the template " & sTemplateFilePath & ". [CommonFunctions.vbs, GetMailContentFromTemplate()]")
	End if
	
	GetMailContentFromTemplate = sMessage
	Set oFSO = nothing
End Function 

'--------------------------------------------------
' NAME: ExtractValueFromFullName
' PURPOSE: Retrieve the various parameter values from the Full Name
'--------------------------------------------------
Function ExtractValueFromFullName(sFName, sParam)
	Dim iStartPos, iEndPos, sValue

	iStartPos = InStr(sFName, sParam)
	If iStartPos > 0 Then
		iStartPos = iStartPos + Len(sParam)
		'Jacob Holcomb - Nov 28, 2007 - If URL field, look for final parenthesis nearest to the beginning, otherwise look for final parenthesis near the end of the string
		If "(URL=" = sParam Then
			iEndPos = InStr(iStartPos, sFName, ")")
		Else
			iEndPos = InStrRev(sFName, ")")
		End If
		sValue = Mid(sFName, iStartPos, iEndPos - iStartPos)
	End If
	ExtractValueFromFullName = sValue
End Function

'--------------------------------------------------
' NAME: GetLinkedURL
' PURPOSE: Retrieve hyperlinked url path
'--------------------------------------------------
Function GetLinkedURL(sURL)
	Dim iPos
	if ( sURL <> "" ) then
		iPos = InStr( 1, sURL, "?" )
		if ( iPos >  0 ) then
			GetLinkedURL = "<a href='" & sURL & "'>" & Left( sURL, iPos-1 ) & 	"</a>" & vbCrLf
		Else
			GetLinkedURL = "<a href='" & sURL & "'>" & sURL & "</a>" & vbCrLf
		end if
	end if
End Function


'--------------------------------------------------
' NAME: GetConfigPath
' PURPOSE: Returns the physical path of parent folder
'--------------------------------------------------
Function GetConfigPath()
	Dim sPath: sPath = WScript.ScriptFullName
	Dim iPos: iPos = InStrRev(sPath, "\")

	sPath = Left(sPath, iPos - 1)
	sPath = Left(sPath, InStrRev(sPath, "\"))
	GetConfigPath = sPath
End Function

'--------------------------------------------------
' NAME: GetVariableValue
' PURPOSE: Retrives the value of variable, which has been passed as argument to the function, from web.config
'--------------------------------------------------
Function GetVariableValue(strVariable)

	Dim		strCfgFileContent
	Dim		FSo, FileObj
	Dim		RegExprObj, ExprMatches, ExprMatch
	Dim		nStartPos, nEndPos

	'Create a file system object
	Set FSo = CreateObject("Scripting.FileSystemObject")
	'Open the configuration file and get a File object
	Set FileObj = FSo.OpenTextFile(GetConfigPath() & "web.config", 1)
	If Err.Number <> 0 Then
		Call Log("ERROR: Failed to load the web.config(" & GetConfigPath() & ") file [" & err.Number & "] " & err.Description & ". [CommonFunctions.vbs, GetMailContentFromTemplate()]")
		Call WScript.Quit(-1)
	End If

	'Read all the contents of the file and store in a string
	strCfgFileContent = FileObj.ReadAll

	'Close the file object
	FileObj.Close

	'Set file objects to Nothing
	Set FileObj = Nothing
	Set FSo = Nothing

	'Create a regular expression object
	Set RegExprObj = New RegExp
	RegExprObj.IgnoreCase = True
	RegExprObj.Global = True
	RegExprObj.Pattern = "\<add\s+key\=\""" + strVariable + "\""\s+value\=\""([^\""]*)\""\s*\/\>"

	'Retrieve the key value according to regular expression pattern
	Set ExprMatches = RegExprObj.Execute(strCfgFileContent)
	'If matching count is greater than one then log error and return empty value (duplicate key-value pair)
	If (ExprMatches.count <> 1) Then
		GetVariableValue = ""
		Exit Function
	End If

	'Get the matching expression from collection
	Set ExprMatch = ExprMatches(0)

	' Since VBScript does not have the handy $1 member that Javascript does,
	' we will do a REPLACE on the string to get the subexpression $1.
	Dim strMatchElement : strMatchElement = ExprMatch.Value
	Dim strResult : strResult = RegExprObj.Replace( strMatchElement, "$1" )

	Set ExprMatches = Nothing
	Set ExprMatch   = Nothing

	GetVariableValue = strResult
End Function


Function encode64( byVal strIn )

	Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
	For n = 1 To Len( strIn ) Step 3
		c1 = Asc( Mid( strIn, n, 1 ) )
		c2 = Asc( Mid( strIn, n + 1, 1 ) + Chr(0) )
		c3 = Asc( Mid( strIn, n + 2, 1 ) + Chr(0) )
		w1 = Int( c1 / 4 ) : w2 = ( c1 And 3 ) * 16 + Int( c2 / 16 )
		If Len( strIn ) >= n + 1 Then
			w3 = ( c2 And 15 ) * 4 + Int( c3 / 64 )
		Else
			w3 = -1
		End If
		If Len( strIn ) >= n + 2 Then
			w4 = c3 And 63
		Else
			w4 = -1
		End If
		strOut = strOut + mimeencode( w1 ) + mimeencode( w2 ) + _
				  mimeencode( w3 ) + mimeencode( w4 )
	Next
	encode64 = strOut
End Function

Function mimeencode( byVal intIn )
    Dim Base64Chars
	Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	If intIn >= 0 Then
		mimeencode = Mid( Base64Chars, intIn + 1, 1 )
	Else
		mimeencode = "="
	End If
End Function

'--------------------------------------------------
' NAME: mimedecode
' PURPOSE: Decode string from Base64
'--------------------------------------------------
Function mimedecode( byVal strIn )
	' Jacob Holcomb - Nov 28, 2007 - Moved Base64Chars inside mimedecode function
	 Dim Base64Chars
	Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	
	If 0 = Len( strIn ) Then
		mimedecode = -1 : Exit Function
	Else
		mimedecode = InStr( Base64Chars, strIn ) - 1
	End If
End Function

Function decode64( byVal strIn )
	Dim w1, w2, w3, w4, n, strOut
	For n = 1 To Len( strIn ) Step 4
		w1 = mimedecode( Mid( strIn, n, 1 ) )
		w2 = mimedecode( Mid( strIn, n + 1, 1 ) )
		w3 = mimedecode( Mid( strIn, n + 2, 1 ) )
		w4 = mimedecode( Mid( strIn, n + 3, 1 ) )
		If w2 >= 0 Then _
			strOut = strOut + _
				Chr( ( ( w1 * 4 + Int( w2 / 16 ) ) And 255 ) )
		If w3 >= 0 Then _
			strOut = strOut + _
				Chr( ( ( w2 * 16 + Int( w3 / 4 ) ) And 255 ) )
		If w4 >= 0 Then _
			strOut = strOut + _
				Chr( ( ( w3 * 64 + w4 ) And 255 ) )
	Next
	decode64 = strOut
End Function

'----------------------------------------------------------------
' NAME:		HideFile
' PURPOSE:	Hides file by setting file attributes
'----------------------------------------------------------------
Function HideFile(strFilePath)

	Dim oFile

	If oFSO.FileExists(strFilePath) Then
	
	    Set oFile = oFSO.GetFile(strFilePath)

	    If Not (oFile.Attributes AND 2 ) Then
		    On Error Resume Next
		    oFile.Attributes = oFile.Attributes XOR 2
		    If err.Number <> 0 Then
			    err.Clear
			    On Error GoTO 0
			    HideFile = False
			    Exit Function
		    End If
	    End If
	    
	Else 
	    Call Log("ERROR: File does not exist: " & strFilePath & ".  Unable to hide file. [CommonFunctions.vbs, GetMailContentFromTemplate()]") 
	End If

	HideFile = True  
	
End Function
