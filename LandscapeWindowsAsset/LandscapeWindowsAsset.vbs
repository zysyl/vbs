'=====================================================================
' NAME:		LandscapeWindowsAsset.vbs
'
' PURPOSE:	Contains functions to query local WMI classes to extract
'			an audit of current hardware and software
'
' USEAGE: 	Assumed this script will execute locally on the target host
'			either as part of a startup or shutdown script.
'
' NOTE:
'
' MODIFIED: (1) 29-May-2011 Version 1.0.0: Initial Release
'
'=====================================================================	

Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Dim wshShell : Set wshShell = CreateObject("Wscript.Shell")
Dim glb_LogFile
Dim glb_logFileExists : glb_logFileExists = False 

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Call PreFlightChecks()
Call getHostname()
Call getOperatingSystem()
Call getADSite()
Call getSerial()
Call getModel()
Call getProcessor()
Call getHardDisks()
Call getMemory()
Call getMac()
Call getIp()
' Running as logon so not executing under system context disable
' this line so as not to generate errors.
'Call getLocalAdmins()
Call getInstalledComponents()


'---------------------------------------------------------------------
' NAME: PreFlightCheck
' PURPOSE:  Run some checks and setup the runtime environment
' @param    
'---------------------------------------------------------------------
Private Sub PreFlightChecks

	On Error Resume Next					
		Dim tmp : tmp = wshShell.ExpandEnvironmentStrings("%TEMP%")		
		Dim host : host = wshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
		glb_LogFile = tmp & "\" & host & "-AUDIT.txt"
		
		Call CreateLocalLogFile(glb_LogFile)
		
		If Err.Number <> 0 Then
			WScript.Quit(-1)
		End If
			
	Err.Clear

End Sub


'---------------------------------------------------------------------
' NAME: getLocalAdmins
' PURPOSE:  Enumerate users in local Administrators group
' @param    
'---------------------------------------------------------------------
Private Sub getLocalAdmins

	Set objGroup = GetObject("WinNT://./Administrators,group")
		On Error Resume Next
		
		Dim i : i = 1
		For Each objUser In objGroup.Members
			Call WriteLogEntry(glb_LogFile,"ADM (" & i & "): " & objUser.Name)
			i = i + 1
		Next
				
		On Error GoTo 0

End Sub


'---------------------------------------------------------------------
' NAME: getADSite
' PURPOSE:  Retrieve AD Site from NLTest
' @param    
'---------------------------------------------------------------------
Private Sub getADSite()		
	Dim wshShell : Set wshShell = CreateObject("WScript.Shell")		
	Dim objNltestOutput : Set objNltestOutput = wshShell.Exec ("cmd /c nltest /dsgetdc:" & domain)
	
	If objNltestOutput.StdOut.AtEndOfStream Then
	 hasNetConn = False
	Else
	 Do Until objNltestOutput.StdOut.AtEndOfStream
		line = objNltestoutput.StdOut.readline
		If InStr(line, "ERROR_NO_SUCH_DOMAIN") Then			
			Exit Do				
		End If
		If InStr(line, "Our Site Name:") > 0 Then
			Call WriteLogEntry(glb_LogFile, Replace(line, "Our Site Name:", "AD Site:") )						
		End If		
	 Loop	
	End If		
End Sub


'---------------------------------------------------------------------
' NAME: getHostname
' PURPOSE:  Retrieve LDAP path from ADSystemInfo
' @param    
'---------------------------------------------------------------------
Private Sub getHostname
	On Error Resume Next 	
		Dim objSysInfo 		: Set objSysInfo = CreateObject("ADSystemInfo")
		'WScript.Echo "Hostname: " & objSysInfo.ComputerName
		Call WriteLogEntry(glb_LogFile, "Hostname: " & objSysInfo.ComputerName)
		If Err.Number <> 0 Then
			'WScript.Echo "Hostname: ERROR"
			Call WriteLogEntry(glb_LogFile,"Hostname: ERROR")
		End If
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: getOperatingSystem
' PURPOSE:  Get Windows OS version from Win32_OperatingSystem
' @param    
'---------------------------------------------------------------------
Private Sub getOperatingSystem
	On Error Resume Next
	
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
		    
		For Each objOperatingSystem in colOperatingSystems
			Call WriteLogEntry(glb_LogFile, "OS: " & objOperatingSystem.Caption & " SP" & objOperatingSystem.ServicePackMajorVersion  _
	        						& "." & objOperatingSystem.ServicePackMinorVersion )		
		Next
		
		If Err.Number <> 0 Then
			Call WriteLogEntry(glb_LogFile, "OS: ERROR")
		End If		
	Err.Clear

End Sub


'---------------------------------------------------------------------
' NAME: getSerial
' PURPOSE:  Retrieve the serial number from Win32_BIOS
' @param    
'---------------------------------------------------------------------
Private Sub getSerial
	On Error Resume Next
	
		Dim objBios : Set objBios = objWMIService.ExecQuery("Select * FROM Win32_BIOS")
		
		If Err.Number <> 0 Then
			Call WriteLogEntry(glb_LogFile, "Serial Number: ERROR")			
		End If
		
		Dim bios
		For Each bios In objBios		
			Call WriteLogEntry(glb_LogFile, "Serial Number: " & bios.serialNumber)			
		Next	
		
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: getProcessor
' PURPOSE:  Retrieves socket count and core count for each CPU from
'			Win32_Processor
' @param    
'---------------------------------------------------------------------
Private Sub getProcessor
	On Error Resume Next
		
		Dim objCpu : Set objCpu = objWMIService.ExecQuery("Select * FROM Win32_Processor")
		
		If Err.Number <> 0 Then
			Call WriteLogEntry(glb_LogFile, "CPU: ERROR")			
		End If
		
		Dim cpu
		Dim i : i = 1
		For Each cpu In objCpu					
			Call WriteLogEntry(glb_LogFile, "CPU (" & i & "): " & Round(cpu.MaxClockSpeed /1000) &  "GHz" )			
			i = i + 1
		Next	
		
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: getModel
' PURPOSE:  Retrieve Model Manufacturer and number of CPU Sockets from
'			Win32_ComputerSystem 
' @param    
'---------------------------------------------------------------------
Private Sub getModel	
	On Error Resume Next
		Dim objCompSystem	: Set objCompSystem = objWMIService.ExecQuery("Select * FROM Win32_ComputerSystem")
		
		If Err.Number <> 0 Then
			Call WriteLogEntry(glb_LogFile,"Model: ERROR")
			Call WriteLogEntry(glb_LogFile,"Manufacturer: ERROR")
		End If
		
		Dim comp
		For Each comp In objCompSystem
			Call WriteLogEntry(glb_LogFile, "Model: " & rtrim(comp.model) )
			Call WriteLogEntry(glb_LogFile,"Manufacturer: " & rtrim(comp.Manufacturer) )
		Next
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: getHardDisks
' PURPOSE:  Retrieve disk size in GB for each disk from Win32_DiskDrive
' @param    
'---------------------------------------------------------------------
Private Sub getHardDisks
	On Error Resume Next
	
	Set diskCollection = objWMIService.ExecQuery("Select * From Win32_DiskDrive")	
	If Err.Number <> 0 Then
		Call WriteLogEntry(glb_LogFile, "HDD: ERROR")
	End If
	
	Dim i : i = 1	
	For Each objDisk In diskCollection
		Call WriteLogEntry(glb_LogFile, "HDD (" & i & "): " &  Round(objDisk.size /1024 /1024 /1024) & "GB" &  vbTab & objDisk.Manufacturer & vbTab & objDisk.Model )
		i = i + 1
	Next	
	
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: GetMemberOfByGUID
' PURPOSE:  Retrieve amount of physical Memory in GB from Win32_PhysicalMemory 
' @param    GUID	The unique ID of a group in AD
'---------------------------------------------------------------------
Private Sub getMemory
	On Error Resume Next
	
		Set memoryCollection = objWMIService.ExecQuery("Select * From Win32_PhysicalMemory")
		If Err.Number <> 0 Then
			Call WriteLogEntry(glb_LogFile, "RAM: ERROR")
		End If		
		Dim i : i = 1		
		For Each objMemory In memoryCollection
			Call WriteLogEntry(glb_LogFile, "RAM ("& i & "): " &  Round(objMemory.Capacity / 1024 / 1034 / 1024) & "GB" & vbTab & objMemory.DeviceLocator)
			i = i + 1
		Next	
		
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: getMac
' PURPOSE:  Retrieve the MAC address for NIC from Win32_NetworkAdapter
'			See script body for ignored interfaces
' @param    
'---------------------------------------------------------------------
Private Sub getMac	
	On Error Resume Next
	
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter")		
		If Err.Number <> 0 Then
			Call WriteLogEntry(glb_LogFile, "MAC: ERROR" )
		End If
		
		Dim i : i = 1
		For Each objItem in colItems
			If Not IsNull(objItem.MACAddress) Then
										
				If InStr(objItem.description, "VMware Virtual Ethernet Adapter") Then					' Ignore vmware player/workstation interfaces
				ElseIf InStr(objItem.description, "RAS") Then											' Ignore RAS interfaces
				ElseIf InStr(objItem.description, "Check Point") Then									' Ignore VPN interfaces
				ElseIf InStr(objItem.description, "Packet Scheduler Miniport") Then						' Ignore
				ElseIf InStr(objItem.description, "WAN Miniport") Then									' Ignore PPTP and PPPOE
				ElseIf InStr(objItem.description, "Deterministic Network Enhancer Miniport") Then		' Ignore 
				ElseIf InStr(objItem.description, "VMware Accelerated") Then 							
				'	Wscript.Echo "MAC: " &  objItem.MACAddress  & vbTab  & objitem.Description 
				Else
					Call WriteLogEntry(glb_LogFile, "MAC (" & i & "): " &  objItem.MACAddress  & vbTab & objitem.Name )
					i = i + 1
				End If
					   			   		
	   		End If   		
		Next
	
	Err.Clear
End Sub


'---------------------------------------------------------------------
' NAME: getMac
' PURPOSE:  Retrieve the MAC address for NIC from Win32_NetworkAdapter
'			See script body for ignored interfaces
' @param    
'---------------------------------------------------------------------
Private Sub getIp
	On Error Resume Next
	
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter " _
	       					& "Where NetConnectionID='Local Area Connection'")
	
		For Each objItem in colItems
	   		strMACAddress = objItem.MACAddress
		Next
		C
		Set colItems = objWMIService.ExecQuery ("Select * From Win32_NetworkAdapterConfiguration")
		If Err.Number <> 0 Then
			all WriteLogEntry(glb_LogFile, "IP: ERROR")
		End If		
		Dim i : i = 1
		For Each objItem in colItems
		    If objItem.MACAddress = strMACAddress Then
		    	If IsArray(objItem.IPAddress) Then		
		    		    		 
			        For Each strIPAddress In objItem.IPAddress		           		           
			           If Instr(strIPAddress,".") Then
			           		all WriteLogEntry(glb_LogFile, "IP (" & i & "): " & strIPAddress)
			           		i = i + 1
			           End If		           
		        	Next 
		        Else
		        	all WriteLogEntry(glb_LogFile, "IP (" & i & "): " & objItem.IPAddress)
		        End If		        
		    End If
		Next
	Err.Clear
End Sub

'---------------------------------------------------------------------
' NAME: getMSHotfixHistory
' PURPOSE:  Retrieve MS hotfix history from local Windows Update service.
' @param    
'---------------------------------------------------------------------
Private Sub getMSHotfixHistory
	Set objSession = CreateObject("Microsoft.Update.Session")
	Set objSearcher = objSession.CreateUpdateSearcher
	intHistoryCount = objSearcher.GetTotalHistoryCount
	
	Set colHistory = objSearcher.QueryHistory(1, intHistoryCount)
	WScript.Echo "Patching"
	For Each objEntry in colHistory
	    Wscript.Echo "Operation: " & objEntry.Operation
	    Wscript.Echo "Result code: " & objEntry.ResultCode	 
	    Wscript.Echo "Date: " & objEntry.Date
	    Wscript.Echo "Title: " & objEntry.Title
	    Wscript.Echo "Client application ID: " & objEntry.ClientApplicationID
	Next
	
End Sub


'---------------------------------------------------------------------
' NAME: getInstalledComponents
' PURPOSE:  Retrieve a list of installed software from Win32_Product.
'			Fallback to registry if Win32_Product fails.
' @param    
'---------------------------------------------------------------------
Private Sub getInstalledComponents
' 	Set productCollection = objWMIService.ExecQuery("Select Name From Win32_Product")		
' 	WScript.Echo "Installed Components"
' 	On Error Resume Next
' 	
' 		For Each objProduct in productCollection
' 			If Not IsNull( objProduct) Then
' 				Wscript.Echo "SFT: " &  objProduct.Name     		
' 			End If		
' 		Next
' 	
' 	If Err.Number <> 0 Then
' 		WScript.Echo "WMI ERROR:- resorting to the registry"
' 		WScript.Echo "ACTUAL ERROR: " & Err.Description
' 		WScript.Echo "========================================="		
		Const uninstallPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
		Dim installedApps
		Dim oReg : Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		oReg.EnumKey HKEY_LOCAL_MACHINE, uninstallPath, installedApps
		Dim i : i = 1
		For Each regProduct In installedApps
			Dim tmpProduct 
			oReg.GetStringValue HKEY_LOCAL_MACHINE, uninstallPath & "\" & regProduct, "Displayname", tmpProduct
						
			' Filter on the following if we are querying the registry
			' Security Update, Hotfix, Update 
			If Not tmpProduct = "" Then		
				
				If InStr(tmpProduct, "Security Update") > 0 Then
				ElseIf InStr(tmpProduct, "Hotfix") > 0 Then
				ElseIf InStr(tmpProduct, "Update") > 0 Then
				Else
					call WriteLogEntry(glb_LogFile, "SFT (" & i & "): " & tmpProduct)
					i = i + 1		
				End If
				
			End If			
		Next		
'	End If
	
End Sub





' NAME: CreateLocalLogFile
' PURPOSE:  Create a new logfile
' @param    logFileName		The full path to the log file
'----------------------------------------------
Sub CreateLocalLogFile(byval logfile)

	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")					
			
	If objFso.FileExists(logfile) Then	
		glb_logFileExists = True		
	Else
		On Error Resume Next
		
			Set localLogFile = objFso.CreateTextFile(logfile,True)		
			If Err.Number <> 0 Then
				glb_logFileExists = False
			Else
				glb_logFileExists = True				
			End If					

		On Error GoTo 0
		Err.Clear
		
	End If
	
	Set objFso = Nothing
	Set shell = Nothing
	
End Sub



'----------------------------------------------
' NAME: WriteLogEntry
' PURPOSE:  Log a message into a logfile
' @param    logFileName		The full path to the log file
' @param	entry			The actual line of text to log
'----------------------------------------------
Sub WriteLogEntry(ByVal logFile, ByVal entry)
	
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
				
	If glb_logFileExists = True Then
		Set txtFile = objFso.OpenTextFile(logFile,8,True)
		txtFile.WriteLine entry		
		txtFile.Close
		'Set txtFile = Nothing		
	End If		
End Sub