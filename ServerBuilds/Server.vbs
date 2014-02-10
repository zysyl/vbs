Class Server

Private m_serial
Private m_model
Private m_macAddresses()


Public Property Get MACAddresses
	MACAddresses = m_macAddresses
End Property

Public Property Get Model
	Model = m_model
End Property

Public Property Get Manufacturer
	Manufacturer = m_manufacturer
End Property

Public Property Get SerialNumber
	SerialNumber = m_serial
End Property


Private Sub setSerial
		Dim objBios : Set objBios = objWMIService.ExecQuery("Select * FROM Win32_BIOS")
		Dim bios
		For Each bios In objBios
			m_serial = bios.serialNumber
		Next	
End Sub

Private Sub setModel
	Dim objCompSystem	: Set objCompSystem = objWMIService.ExecQuery("Select * FROM Win32_ComputerSystem")
	Dim comp
	For Each comp In objCompSystem
		m_model = rtrim(comp.model)
	Next
End Sub

Private Sub setMac
	' Get MAC address of Local Area Connection
	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter Where NetConnectionID='Local Area Connection'")		
	For Each objItem in colItems
   		Redim Preserve m_macAddresses( objItem.MACAddress   )
	Next
End Sub

End Class

