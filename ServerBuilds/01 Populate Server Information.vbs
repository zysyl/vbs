' Tag the AD machine object for this computer with information like chassis model etc.
Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const LOG_FILE_PATH = "C:\Windows\"	
Const LOG_FILE_NAME = "sbet-startup.txt"

Dim objLocalMachine : Set objLocalMachine = New Machine
Dim objSysInfo 		: Set objSysInfo = CreateObject("ADSystemInfo")
Dim objADMachine    : Set objADMachine = GetObject("LDAP://" & objSysInfo.ComputerName)
Dim logFileExists 	: logFileExists = False

Call CreateLocalLogFile()
Call WriteLogEntry("[INFO]","Starting " & WScript.ScriptName)
Call WriteLogEntry("[INFO]","Machine Name " & objSysInfo.ComputerName)

'Call WriteADDescription()
'Call WriteADSerial()
'Call WriteADIpHostNumber()

Call WriteLogEntry("[INFO]","Stopping " & WScript.ScriptName)

' ==== End of Script ===

Sub WriteADDescription()

	On Error Resume Next
		objADMachine.PutEx ADS_PROPERTY_CLEAR, "description", 0
		objADMachine.SetInfo		
		objADMachine.Put "description", objLocalMachine.getChassis & "," & _
									objLocalMachine.getModel  & "," & _
									 objSysInfo.SiteName & ",SWD" & _
									  objLocalMachine.getBuildVersion												 									  
		objADMachine.SetInfo
		
		If Err.Number <> 0 Then
			Call WriteLogEntry ("[ERROR]", "Error writing machine description attribute the actual error was. " & Err.Description )						
		End If
				
	On Error GoTo 0
	Err.clear

End Sub

Sub WriteADSerial( )
	On Error Resume Next	
		objADMachine.PutEx ADS_PROPERTY_CLEAR, "serialNumber", 0
 		objADMachine.SetInfo
		objADMachine.PutEx ADS_PROPERTY_APPEND, "serialNumber", Array(objLocalMachine.getSerial)
		objADMachine.SetInfo				
		
		If Err.Number <> 0 Then
			Call WriteLogEntry( "[ERROR]", "Error writing Serial number attribute the actual error was. " & Err.Description )
		End If		
		
	On Error GoTo 0
	Err.clear
End Sub

Sub WriteADIpHostNumber()
	On Error Resume Next
		objADMachine.PutEx ADS_PROPERTY_CLEAR, "ipHostNumber", 0
		objADMachine.SetInfo
		objADMachine.PutEx ADS_PROPERTY_APPEND, "ipHostNumber", Array(objLocalMachine.getIp)
		objADMachine.SetInfo
		
		If Err.Number <> 0 Then
			Call WriteLogEntry( "[ERROR]","Problem encountered whilst writing ipHostNumber attribute, the error message was. " & Err.Description )
		Else
			Call WriteLogEntry( "[INFO]","IP Address is " & objLocalMachine.getIp & " written to ipHostNumber")
		End If
				
	On Error GoTo 0
	Err.Clear
End Sub

Sub CreateLocalLogFile()

	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
			
	Dim fullLogFilePath : fullLogFilePath = LOG_FILE_PATH & "\" & LOG_FILE_NAME	
	Dim localLogFile
			
	If objFso.FileExists(fullLogFilePath) Then	
		logFileExists = True		
	Else
		On Error Resume Next
		
			Set localLogFile = objFso.CreateTextFile(fullLogFilePath,True)		
			If Err.Number <> 0 Then
				logFileExists = False
			Else
				logFileExists = True				
			End If					

		On Error GoTo 0
		Err.Clear
		
	End If
	
	Set objFso = Nothing
	Set shell = Nothing
	
End Sub

Sub WriteLogEntry(ByVal msgType, ByVal entry)
	
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
	Dim logFile	
	Dim fullLogFilePath : fullLogFilePath = LOG_FILE_PATH & "\" & LOG_FILE_NAME	
			
	If logFileExists = True Then
		Set logFile = objFso.OpenTextFile(fullLogFilePath,8,True)
		logFile.WriteLine Now & vbTab & msgType & vbTab & entry
		logFile.Close
		Set logFile = Nothing		
	End If		
End Sub

Class Machine
' #region Persistent fold region == Instance Variable

	Private domain
	Private site
	Private model
	Private architecture
	Private chassis
	Private localDC
	Private serial	
	Private prefix
	Private ip
	Private hasNetConn	
	
	Dim wshShell
	Dim objNltestOutput 
	Dim objWMIService
	Dim objSysInfo
' #endregion
 
	' Constructor
	Public Sub class_initialize
		domain = "UNKNOWN"
		site = "UNKNOWN"		
		model = "UNKNOWN"
		architecture = "UNKNOWN"
		chassis = "UNKNOWN"
		localDC = "UNKNOWN"
		serial = "UNKNOWN"
		ip = "UNKNOWN"
		hasNetConn = True		
		
		Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
		Set wshShell = CreateObject("Wscript.Shell")
		Set objSysInfo = CreateObject("ADSystemInfo")
		Call InitialiseMachine()
		
	End Sub

	' Destructor
	Public Sub class_terminate
		set objWMIService = Nothing
		Set wshShell = Nothing
		Set objNltestOutput = Nothing	
		Set objSysInfo = Nothing	
	End Sub

	' Utility methods
	Private Sub InitialiseMachine()
	
		' Populate the properties of this object based on the output from NLTest
		On Error Resume Next
			Call setDomain
			Call ParseNlTestInformation			
			Call setSerial					
			Call setModel
			Call setArchitecture
			Call setChassis					
			Call setIp
		Err.Clear
		
	End Sub		
	
	Private Sub ParseNlTestInformation()		
		Set objNltestOutput = wshShell.Exec ("cmd /c nltest /dsgetdc:" & domain)
		
		If objNltestOutput.StdOut.AtEndOfStream Then
		 hasNetConn = False
		Else
		 Do Until objNltestOutput.StdOut.AtEndOfStream
			line = objNltestoutput.StdOut.readline
			If InStr(line, "ERROR_NO_SUCH_DOMAIN") Then
				hasNetConn = False
				Exit Do				
			End If
			If InStr(line, "Our Site Name:") > 0 Then
				Call setSite (line)				
			End If
			If InStr(line, "DC:") > 0 Then
				Call setLocalDC(line)
			End If				
		 Loop	
		End If		
	End Sub
	
	' Setter methods	
	Private Sub setDomain()
		domain = objSysInfo.DomainShortName
	End Sub
		
	Private Sub setSite(nltestLine)	
		Dim tmpStr : tmpStr = Split(nltestLine,":")(1)	
		site = trim(tmpStr)
	End Sub

	Private Sub setLocalDC(nltestLine)
		localDC = trim(Replace(Split(nltestLine,":")(1),"\\",""))
	End Sub

	Private Sub setSerial
		Dim objBios : Set objBios = objWMIService.ExecQuery("Select * FROM Win32_BIOS")
		Dim bios
		For Each bios In objBios
			serial = bios.serialNumber
		Next	
	End Sub

	Private Sub setModel
		Dim objCompSystem	: Set objCompSystem = objWMIService.ExecQuery("Select * FROM Win32_ComputerSystem")
		Dim comp
		For Each comp In objCompSystem
			model = rtrim(comp.model)
		Next
	End Sub

	Private Sub setArchitecture
		architecture = wshshell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
	End Sub

	Private Sub setChassis
		' determine if we are a desktop, laptop or virtual Machine
		' additional checks to determine if we are a server chassis
		Dim iBattery			: Set iBattery = GetObject("winmgmts:\\.\root\CIMV2").InstancesOf ("Win32_Battery")
		Dim iPortableBattery	: Set iPortableBattery = GetObject("winmgmts:\\.\root\CIMV2").InstancesOf ("Win32_PortableBattery")
		Dim iPCMCIAController	: Set iPCMCIAController = GetObject("winmgmts:\\.\root\CIMV2").InstancesOf ("Win32_PCMCIAController")
		Dim result				: result = iBattery.count + iPortableBattery.count + iPCMCIAController.count	
		
		If result <> 0 Then
			chassis = "Laptop"
		Else
			chassis = "Desktop"
		End If
		
		If InStr(serial, "VMware") Then
			chassis = "Virtual"
		End If	
	End Sub 

	Private Sub setIp
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter " _
        					& "Where NetConnectionID='Local Area Connection'")

		For Each objItem in colItems
    		strMACAddress = objItem.MACAddress
		Next
		
		Set colItems = objWMIService.ExecQuery ("Select * From Win32_NetworkAdapterConfiguration")

		For Each objItem in colItems
		    If objItem.MACAddress = strMACAddress Then
		    	If IsArray(objItem.IPAddress) Then		    		 
			        For Each strIPAddress In objItem.IPAddress		           		           
			           If Instr(strIPAddress,".") Then
			           		ip = strIPAddress
			           End If		           
		        	Next 
		        Else
		        	ip = objItem.IPAddress
		        End If		
		        Exit For
		    End If
		Next
	End Sub


	' Accessor functions
	Public Function getSite
		getSite = site
	End Function
	
	Public Function getSerial
		getSerial = serial
	End Function
	
	Public Function getModel
		getModel = model
	End Function
	
	Public Function getBuildVersion
		getBuildVersion = wshShell.RegRead("HKLM\System\Sportingbet\Build\version")	
	End Function	
	
	Public Function getArchitecture
		getArchitecture = architecture	
	End Function
	
	Public Function getChassis
		getChassis = chassis
	End Function
	
	Public Function getDomain
		getDnsZone = domain
	End Function
	
	Public Function getLocalDC
		getLocalDC = localDC
	End Function
			
	Public Function getIp
		getIp = ip
	End Function
	
	Public Function IsOnSbetNetwork			
		IsOnSbetNetwork = hasNetConn			
	End Function
		
End Class