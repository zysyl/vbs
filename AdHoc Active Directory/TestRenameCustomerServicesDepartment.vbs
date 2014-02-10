Const ADS_PROPERTY_CLEAR = 1

Dim guid : guid = "F8F84422061EBF4F90E46AFDB6CA34F3"
	
' Bind to user account
Dim objUserAccount : Set objUserAccount = GetObject("LDAP://<GUID=" & guid & ">")				

WScript.Echo "  Old: " & objUserAccount.Department
		
' Clear out old description field entry
objUserAccount.PutEx ADS_PROPERTY_CLEAR, "department", 0	

' Write date disabled in description
objUserAccount.Put "department", "IT"
	
' Save Changes
objUserAccount.SetInfo

WScript.Echo "  New: " & objUserAccount.Department