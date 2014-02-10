Dim glb_ObjADConnection, glb_ObjADCommand, glb_objADRecordset

Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_UF_ACCOUNTDISABLE = 2
Const SQL_STATEMENT = "SELECT givenName, sn, samAccountName, msExchHideFromAddressLists, userAccountControl FROM 'LDAP://DC=Sbet-EMEA,DC=ADS' WHERE objectCategory='user' AND objectClass='user' ORDER BY samAccountName"

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

	' Echo to screen the values
	Dim isHidden : isHidden = glb_objADRecordset.Fields("msExchHideFromAddressLists")
	
	If isHidden = "True" Then
		WScript.Echo glb_objADRecordset.Fields("samAccountName") & vbTab &_
			glb_objADRecordset.Fields("givenName") & vbTab &_			
				glb_objADRecordset.Fields("sn") & vbTab &_
					glb_objADRecordset.Fields("msExchHideFromAddressLists") & vbTab &_
						glb_objADRecordset.Fields("userAccountControl")
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