'=====================================================================
' NAME:		commom.vbs
' PURPOSE:	Contains common utility functions shared by scripts
'=====================================================================	

'----------------------------------------------
' NAME: GetComputerModel
' PURPOSE:  Detecs the model of the machine from WMI
' @param    
'----------------------------------------------
Private Function GetComputerModel
	Dim objWMIService   : Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")		
	Dim objCompSystem	: Set objCompSystem = objWMIService.ExecQuery("Select * FROM Win32_ComputerSystem")
	Dim comp, model
	For Each comp In objCompSystem			
		model = rtrim(comp.model)
	Next
	GetComputerModel = model
End Function


'----------------------------------------------
' NAME: GetOSArchitecture
' PURPOSE:  Detecs the Architecture of the machine
' @param    
'----------------------------------------------
Private Function GetOSArchitecture		
	Dim objWMIService   : Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")		
	Dim objCompSystem	: Set objCompSystem = objWMIService.ExecQuery("Select * FROM Win32_OperatingSystem")
	Dim os, arch, tmpArch
	tmpArch = "UN"	
	For Each os In objCompSystem				
		On Error Resume Next
			tmpArch = os.OSArchitecture
			If Err.Number <> 0 Then				
				Dim wshShell : Set wshShell = CreateObject("Wscript.Shell")
				tmpArch = wshshell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")				
			End If			
		Err.Clear		
	Next
	
	Select Case tmpArch			
		Case "32-bit":
			arch = "X86"
		Case "x86"
			arch = "X86"		
		Case "64-bit":
			arch = "X64"
		Case "amd64"
			arch = "X64"
		Case Else:
			arch = "UN"					
	End Select	
	
	GetOSArchitecture = arch
End Function


'----------------------------------------------
' NAME: GetComputerOS
' PURPOSE:  Detecs the model of the machine from WMI
' NOTE:  Assumes Windows Pre-Execution Environment If
'		 SystemDrive is X:
' @param    
'----------------------------------------------
Private Function GetComputerOS

	Dim objWMIService   : Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")	
	Dim objCompOS		: Set objCompOS = objWMIService.ExecQuery("Select * FROM Win32_OperatingSystem")
	Dim os, version
	For Each os In objCompOs
		version = os.Caption
		If version = "" Then
			Dim wshShell 	: Set wshShell = CreateObject("Wscript.Shell")
			Dim systemDrive : systemDrive = wshshell.ExpandEnvironmentStrings("%SystemDrive%")
			If systemDrive = "X:" Then	
				version = "Microsoft Windows Pre-Execution Environment"
			End If			
		End If		
	Next
	
		
	GetComputerOS = version
End Function



'----------------------------------------------
' NAME: CreateLocalLogFile
' PURPOSE:  Create a new logfile
' @param    logFileName		The full path to the log file
'----------------------------------------------
Sub CreateLocalLogFile(ByVal logFileName)	
	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
				
	Dim fullLogFilePath : fullLogFilePath = ".\" & logFileName	
	Dim localLogFile
			
	If objFso.FileExists(fullLogFilePath) Then	
		glb_logFileExists = True		
	Else
		'On Error Resume Next
		
			Set localLogFile = objFso.CreateTextFile(fullLogFilePath,True)	
			If Err.Number = 0 Then
				glb_logFileExists = True
			Else
				glb_logFileExists = False
			End If			
			
		'On Error GoTo 0
		'Err.Clear
		
	End If	
End Sub


'----------------------------------------------
' NAME: Log
' PURPOSE:  Log a message into a logfile
' @param    logFileName		The full path to the log file
' @param	entry			The actual line of text to log
'----------------------------------------------
Sub WriteLogEntry(ByVal logFileName, ByVal entry)
	Dim logFile	
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
	Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")		
	Dim fullLogFilePath : fullLogFilePath = ".\" & logFileName	
			
	If glb_logFileExists = True Then
		Set logFile = objFso.OpenTextFile(fullLogFilePath,8,True)
		logFile.WriteLine entry
		logFile.Close
	End If		
End Sub



'----------------------------------------------
' NAME: RunPCI32
' PURPOSE:  Will run PCI32 with -i and append ouput to a logfile
' @param    logFileName		The full path to the log file
'----------------------------------------------
Sub RunPCI32(Byval logFileName)
	Dim shell : Set shell = CreateObject("Wscript.Shell")	
	shell.run "cmd /c .\PCI32\Pci32.exe -i >> "  & logFileName, 0, True	
End Sub
