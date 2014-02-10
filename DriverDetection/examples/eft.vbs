'--------------------------------------------------
' NAME: GetUserEmailFromEFT
' PURPOSE: Returns a user's email address from EFT
'--------------------------------------------------
Function GetUserEmailFromEFT(EFTServerLocation, EFTPort, EFTAdminUserName, EFTAdminPassword, siteId, settingsLevel, tempAccountName)
    
    Dim SFTPServer
    
    ' Instantiate the SFTPCOM Admin object and connect to the server.
    Set SFTPServer = WScript.CreateObject("SFTPCOMInterface.CIServer")

    On Error Resume Next
    'Connect to EFT
    SFTPServer.Connect EFTServerLocation, EFTPort, EFTAdminUserName, EFTAdminPassword
    
    If Err.Number <> 0 Then
       Call Log("ERROR: Error connecting to '" & EFTServerLocation & ":" & EFTPort & "' -- " & err.Description & " [" & CStr(err.Number) & "]")
       WScript.Quit(255)
    Else
       Call Log("INFO: Successfully connected to '" & EFTServerLocation & "'")
    End If
    
    If SFTPServer <> NULL Then
        Dim Sites, Site
        Set Sites = SFTPServer.Sites
        ' Determine the integer SiteID if a Site Name is given:
        Dim iSiteCount
     
        If (Not isNumeric(siteId)) Then         
            'The siteId is actually the site name
            Call Log("INFO: Attempting to find site with name: " & siteId)
            For iSiteCount = 0 to Sites.Count - 1
                Set Site = Sites.Item(iSiteCount)
                If LCase(Trim(Site.Name)) = LCase(Trim(siteId)) Then
                    Call Log("INFO: Successfully found site: " & siteId)
                    Exit For
                End If
            Next      
        Else
            'siteId is is numeric "id"
            Call SLog("INFO: Loading site with ID: " & siteId)
            Set Site = Sites.SiteByID(CLng(siteId))
        End If
        
        Dim oSettings
        Set oSettings = Site.GetUserSettings(tempAccountName)
        
        'retrieve user email from EFT
        Dim userEmail
        userEmail = oSettings.Email   
         
        If SFTPServer <> "" Then 
            Call Log("INFO: Retrieved user email from EFT: " & userEmail) 
        Else
            Call SLog("ERROR: User email is blank from EFT ") 
        End If
        
    Else 
       Call Log("ERROR: SFTPServer object is NULL") 
       userEmail = ""
    End If
    
    'Close connection to EFT
    SFTPServer.Close
	Set SFTPServer = nothing

    GetUserEmailFromEFT = userEmail   
End Function