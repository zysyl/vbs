'=====================================================================
' NAME:		BuildCompleteMoveComputerObject.vbs
'
' PURPOSE:	Once a desktop, laptop or virtual machine has 
'			completed the build stage, move it to live OU.
'
' MODIFIED: (1) 01/08/2011 Version 1.00: Initial Release 
'=====================================================================
Option Explicit

Const LCL_BUILD_SOURCE = "C:\SWD"
Const LOG_FILE_PATH = "C:\Windows\System32"	
Const LOG_FILE_NAME = "swd-build.txt"
Const BUILD_COMPLETE = "HKLM\System\Sportingbet\build\complete"			' Create this flag in SWDBuildEngine
Dim glb_SysInfo, glb_ADMachine, glb_Chassis, glb_TargetOU, glb_XmlDoc, glb_logFileExists

Call WriteLogEntry("[INFO]","Starting " & WScript.ScriptName)

If RegistryKeyExists(BUILD_COMPLETE) = True Then
		
	Set glb_SysInfo = CreateObject("ADSystemInfo")
	Set glb_ADMachine = GetObject("LDAP://" & glb_SysInfo.ComputerName)
	Set glb_XmlDoc = CreateObject("Microsoft.XMLDom")
	
	glb_Chassis = DetectMachineChassis
	
	Select Case glb_SysInfo.SiteName
	
		Case "IEDublinCLN" :
		
			If glb_Chassis = "Laptop" Then
				glb_TargetOU = "OU=Laptops,OU=Computers,OU=BEC,OU=Ireland,DC=Sbet-EMEA,DC=ADS"
			ElseIf glb_Chassis = "Desktop" Then
				glb_TargetOU = "OU=Workstations,OU=Computers,OU=BEC,OU=Ireland,DC=Sbet-EMEA,DC=ADS"
			ElseIf glb_Chassis = "Virtual" Then
				glb_TargetOU = "OU=Computers,OU=BEC,OU=Ireland,DC=Sbet-EMEA,DC=ADS"
			End If		
					
		Case "UKGuernseySTP" :
			If glb_Chassis = "Laptop" Then
				glb_TargetOU = "OU=Laptops,OU=Computers,OU=STP,OU=Guernsey,DC=Sbet-EMEA,DC=ADS"
			ElseIf glb_Chassis = "Desktop" Then
				glb_TargetOU = "OU=Workstations,OU=Computers,OU=STP,OU=Guernsey,DC=Sbet-EMEA,DC=ADS"
			ElseIf glb_Chassis = "Virtual" Then
				glb_TargetOU = "OU=Computers,OU=STP,OU=Guernsey,DC=Sbet-EMEA,DC=ADS"
			End If	
			
		Case "UKLondonTWH" :
			If glb_Chassis = "Laptop" Then
				glb_TargetOU = "OU=Laptops,OU=Computers,OU=TWH,OU=UK,DC=Sbet-EMEA,DC=ADS"
			ElseIf glb_Chassis = "Desktop" Then
				glb_TargetOU = "OU=Workstations,OU=Computers,OU=TWH,OU=UK,DC=Sbet-EMEA,DC=ADS"
			ElseIf glb_Chassis = "Virtual" Then
				glb_TargetOU = "OU=VM Desktop,OU=Computers,OU=TWH,OU=UK,DC=Sbet-EMEA,DC=ADS"
			End If
	
	End Select
	
	' Log some information 		
	Call WriteLogEntry("[INFO] Target OU", glb_TargetOU )	
	
	' Remove local build source
	Call RemoveLocalBuildSource( LCL_BUILD_SOURCE )
	
	' Last task is to move the computer object
	Call MoveMachine( glb_TargetOU )
	
End If


Call WriteLogEntry("[INFO]","Stopping " & WScript.ScriptName)

'----------------------------------------------
' NAME: DetectMachineChassis
' PURPOSE:  Detects if a client machine is a
'			laptop, desktop or virtual
' @param    target the directory to remove
'----------------------------------------------
Sub RemoveLocalBuildSource(target)
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	
	If fso.FolderExists(target) Then
    	fso.DeleteFolder target
    End If    
    
End Sub


'----------------------------------------------
' NAME: DetectMachineChassis
' PURPOSE:  Detects if a client machine is a
'			laptop, desktop or virtual
' @param    
'----------------------------------------------
Function DetectMachineChassis

Dim iBattery			: Set iBattery = GetObject("winmgmts:\\.\root\CIMV2").InstancesOf ("Win32_Battery")
Dim iPortableBattery	: Set iPortableBattery = GetObject("winmgmts:\\.\root\CIMV2").InstancesOf ("Win32_PortableBattery")
Dim iPCMCIAController	: Set iPCMCIAController = GetObject("winmgmts:\\.\root\CIMV2").InstancesOf ("Win32_PCMCIAController")
Dim objWMIService 		: Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Dim objBios 			: Set objBios = objWMIService.ExecQuery("Select * FROM Win32_BIOS")
Dim result				: result = iBattery.count + iPortableBattery.count + iPCMCIAController.count	
Dim bios, serial

For Each bios In objBios
	serial = bios.serialNumber
Next	
		
If result <> 0 Then
	DetectMachineChassis = "Laptop"
Else
	DetectMachineChassis = "Desktop"
End If

If InStr(serial, "VMware") Then
	DetectMachineChassis = "Virtual"
End If	

End Function



'----------------------------------------------
' NAME: RegistryKeyExists
' PURPOSE:  Detects if a given registry key exists
' @param    regKey the registry key to check for
'----------------------------------------------
Function RegistryKeyExists( ByVal regKey )
	Dim wshShell : Set wshShell = CreateObject("Wscript.Shell")
	On Error Resume Next
	
		wshShell.RegRead(regKey)
		
		If Err.Number <> 0 Then
			RegistryKeyExists = False
		Else
			RegistryKeyExists = True					
		End If
	
	On Error GoTo 0
	Err.Clear

End Function


'----------------------------------------------
' NAME: MoveMachine
' PURPOSE:  Move a computer object in AD
' @param    targetOU the AD OU to move this object
'			to
'----------------------------------------------
Sub MoveMachine(ByVal targetOU)

Dim MoveToOU : Set MoveToOU = GetObject("LDAP://" & targetOU)
On Error Resume Next
	' Trap any access denied type error
	Err.Clear
	InfoMessage "Moving machine object to " & targetOU
	MoveToOU.MoveHere "LDAP://" & objSysInfo.ComputerName, vbNullString		
	intError=Err.Number
	strError=Err.Description
On Error GoTo 0

End Sub


'----------------------------------------------
' NAME: CreateLocalLogFile
' PURPOSE:  Create a new logfile
' @param    logFileName		The full path to the log file
'----------------------------------------------
Sub CreateLocalLogFile()

	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
			
	Dim fullLogFilePath : fullLogFilePath = LOG_FILE_PATH & "\" & LOG_FILE_NAME	
	Dim localLogFile
			
	If objFso.FileExists(fullLogFilePath) Then	
		glb_logFileExists = True		
	Else
		On Error Resume Next
		
			Set localLogFile = objFso.CreateTextFile(fullLogFilePath,True)		
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
' NAME: Log
' PURPOSE:  Log a message into a logfile
' @param    logFileName		The full path to the log file
' @param	entry			The actual line of text to log
'----------------------------------------------
Sub WriteLogEntry(ByVal msgType, ByVal entry)
	
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
	Dim logFile	
	Dim fullLogFilePath : fullLogFilePath = LOG_FILE_PATH & "\" & LOG_FILE_NAME	
			
	If glb_logFileExists = True Then
		Set logFile = objFso.OpenTextFile(fullLogFilePath,8,True)
		logFile.WriteLine Now & vbTab & msgType & vbTab & entry
		logFile.Close
		Set logFile = Nothing		
	End If		
End Sub
