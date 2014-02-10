'=====================================================================
' NAME:		GetMemberOf
'
' PURPOSE:	Contains functions to aid with retrieving the members of
'			an Active Directory Group
'
' USEAGE:	General Purpose to query Active Directory
'
' NOTE:		
'
' MODIFIED: 02-Sep-2011 Version 1.0.0: Initial Release
'
'=====================================================================	

'GAVS Sysadmins
Call GetMemberOfByGUID( "B01B1EBE7AAC9D4E84CE23D45B82D175")

'MOR-Perforce-Administrators
Call GetMemberOfByGUID( "25CCED887381DA458C46A370CD88C387" )

'Role TWH Server Admins
Call GetMemberOfByGUID( "27E30E19DD80A34A8B88799B86EC6518" )

'Role TWH Server RDP Access
Call GetMemberOfByGUID ( "6485CB441182EC4697F43D8FAD18611D" )

' Domain Admins
Call GetMemberOfByGUID ( "A756491F817DAA4A85A2A0060B2E66EC" )
'---------------------------------------------------------------------
' NAME: GetMemberOfByGUID
' PURPOSE:  Retrieve members of an AD group 
' @param    GUID	The unique ID of a group in AD
'---------------------------------------------------------------------
Function GetMemberOfByGUID( ByVal GUID )	
		
	Dim groupMembers, adDomainPath     
	Dim connection : set connection = CreateObject("ADODB.Connection") 
	Dim command : Set command = CreateObject("ADODB.Command")
	
    adDomainPath = "LDAP://<GUID=" + GUID + ">"       
    
    ' Connection Object 
    connection.Provider = "ADsDSOObject"
    'connection.Properties.Item("User ID") = "srv-sbsoe"        
    'connection.Properties.Item("Password") = "Pa55w0rd!"
    'connection.Properties.Item("Encrypt Password") = True
    connection.Open("Active Directory Provider")   
    
    ' Command Object
    command.ActiveConnection = connection	    
    command.CommandText = "select member from '" + adDomainPath + "'"
    command.Properties.Item("Page Size") = 10000
    command.Properties.Item("Timeout") = 30
    command.Properties.Item("SearchScope") = 2
    command.Properties.Item("Cache Results") = True
    
    Set recordSet = command.Execute()    
    
    If recordSet.EOF Or recordSet.BOF Then
    	' the result set is empty    	
    	GetMemberOfByGUID = Null
    Else 
    	' the result set contains entries   	
    	groupMembers = recordSet.Fields("member")
    	For Each member In groupMembers    		
			Call PrintDetails ( member )
		Next		
	End If
	
	' close the connection to AD  	  	
  	connection.Close
  	GetMemberOfByGUID = recordSet 	
   			
End Function	


'---------------------------------------------------------------------
' NAME: CountMemberByGUID
' PURPOSE:  Retrieve how many objects are in a group 
' @param    GUID	The unique ID of a group in AD
'---------------------------------------------------------------------
Function CountMemberByGUID( ByVal GUID )

	Dim groupMembers, adDomainPath, membersCount     
    adDomainPath = "LDAP://<GUID=" + GUID + ">"       
    
    ' Connection Object 
    connection.Provider = "ADsDSOObject"
    connection.Properties.Item("User ID") = "srv-sbsoe"        
    connection.Properties.Item("Password") = "Pa55w0rd!"
    connection.Properties.Item("Encrypt Password") = True
    connection.Open("Active Directory Provider")   
    
    ' Command Object
    command.ActiveConnection = connection	    
    command.CommandText = "select member from '" + adDomainPath + "'"
    command.Properties.Item("Page Size") = 10000
    command.Properties.Item("Timeout") = 30
    command.Properties.Item("SearchScope") = 2
    command.Properties.Item("Cache Results") = True
    
    Set recordSet = command.Execute()    
    
    If recordSet.EOF Or recordSet.BOF Then
    	' the result set is empty    	
    	membersCount = 0
    Else 
    	' the result set contains entries   	    	
    	membersCount = recordSet.Rows.Count
	End If
	
	' close the connection to AD  	  	
  	connection.Close
  	CountMemberByGUID = recordSet 	

End Function



'---------------------------------------------------------------------
' NAME: DisableUserAccount
' PURPOSE:  Disable a user account in AD and write to the
'			objects description field
' @param    guid		The unique id of the account object
'---------------------------------------------------------------------
Function PrintDetails( dn)
	Dim objUser : Set objUser = GetObject("LDAP://" & dn)	
	
	' Spit out to screen
	WScript.Echo objUser.DisplayName & vbTab & objUser.UserAccountControl
					
End Function