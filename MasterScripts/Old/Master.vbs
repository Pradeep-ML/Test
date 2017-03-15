'#################################Master Script####################################################
'Purpose: Kick off Automation Framework and launch all vbscripts in a sequence
'VBScript files call sequence:  KillProcesses_CleanFolders.vbs
'								InstallTrust.vbs
'								LaunchVNC.vbs
'								ExecuteTestSet.vbs
'##################################################################################################

'Declare all variables
Dim strCurrentPath
Dim blnParentFolderFound
Dim strCleanupPath

'Get current directory
Set objWScript = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentPath = WScript.ScriptFullName

'Loop until "MobileLabs Automation Framework" folder is found
blnParentFolderFound = True
Do While Replace(Split(strCurrentPath,"\")(UBound(Split(strCurrentPath,"\"))), " ", "") <> "MobileLabsAutomationFramework"
	strCurrentPath = objFSO.GetParentFolderName(strCurrentPath)
	'Exit if reaches the system drive
	If InStr(1, strCurrentPath, "\") = 0 Then
		blnParentFolderFound = False
		Exit Do
	End If
Loop

'Define the path of the Root Folder: <MobileLabs Automation Framework>
If blnParentFolderFound Then
	If Right(strCurrentPath,1) <> "\" Then
		strCurrentPath = strCurrentPath & "\"
	End If
	If Replace(objFSO.GetFolder(strCurrentPath).Name, " ", "") <> "MobileLabsAutomationFramework" Then
		MsgBox "MobileLabs Automation Framework folder was not found!"
		WScript.Quit
	End If
Else
	MsgBox "Error: Master.vbs file is being executed from a wrong location: " & objWScript.CurrentDirectory
	WScript.Quit
End If

'Execute all VBS files in a defined order:

'Launch the cleanup file
MsgBox "Start cleanup!"
strCleanupPath = strCurrentPath & "Cleanup\KillProcesses_CleanFolders.vbs"
objWScript.Run Chr(34) & strCleanupPath & Chr(34)
WaitForVBS 1000		'Wait for the Cleanup script to finish
MsgBox "Cleanup done"

'Download and install Trust
MsgBox "Download Build if applicable and install"
'strInstallationVBS = strCurrentPath & "Environment Setup\InstallTrust.vbs"
'bjWScript.Run Chr(34) & strInstallationVBS & Chr(34)
'WaitForVBS	60000	'Wait for the Installation script to finish
MsgBox "Installation Done"

'Create Test Set
MsgBox "Create Test Set"
strCreateTestSetPath = strCurrentPath & "Master\CreateTestSet.vbs"
objWScript.Run Chr(34) & strCreateTestSetPath & Chr(34)
WaitForVBS 1000		'Wait for the Create Testset script to finish
MsgBox "Test Set created"

'Setup Environment
MsgBox "Setting up environment"
strLaunchDevicePath = strCurrentPath & "Environment Setup\LaunchDeviceViewer.vbs"
objWScript.Run Chr(34) & strLaunchDevicePath & Chr(34)
WaitForVBS 2000		'Wait for the Environment script to finish
MsgBox "Environment Setup done"

'Execute Scripts
MsgBox "Start execution of tests"
strExecutionPath = strCurrentPath & "Master\ExecuteTestSet.vbs"
objWScript.Run Chr(34) & strExecutionPath & Chr(34)
WaitForVBS 10000	'Wait for the Execution script to finish
MsgBox "Test Set executed"

Set objWScript = Nothing
Set objFSO = Nothing

MsgBox "Test Complete. Please proceed to analyze and log defects..!!"

'This sub waits until there is only one wscript.exe process running
Sub WaitForVBS(intTimeOut)
	'Create an object of WMI
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 

	'Executing query to get the list of all wscript.exe processes
	Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process where name like 'wscript.exe'")
	
	'Get the count and if greater than 1 then do recursion
	If colProcess.Count > 1 Then
		Set colProcess = Nothing
		Set objWMIService = Nothing
		WScript.Sleep intTimeOut
		WaitForVBS intTimeOut
	End If
End Sub