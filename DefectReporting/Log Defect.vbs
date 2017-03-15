'Declare all variables
Dim strCurrentPath
Dim blnParentFolderFound
Dim strCleanupPath

'Get current directory
Set objWScript = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentPath = objWScript.CurrentDirectory

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
End If

'Check if any excel file exists in the strCurrentPath
Dim Test_path, filesys
Test_path  = strCurrentPath & "DefectReporting\Log Defect"

	blnFilesFound = False
	FilesPath = strCurrentPath & "DefectReporting"
	Set objFiles = ObjFSO.GetFolder(FilesPath).Files
	for each objFile  in objFiles
	 If LCase(Right(objFile.Name, 4)) = "xlsx" OR LCase(Right(objFile.Name, 3)) = "xls" Then
		blnFilesFound = True
	 End If 
	Next

	Set ObjFSO=Nothing
	Set DelFilesObj=Nothing

set filesys=CreateObject("Scripting.FileSystemObject")

If  blnFilesFound Then
	Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
	Dim qtTest 'As QuickTest.Test ' Declare a Test object variable
	Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
	'Load required Add-ins
	qtApp.SetActiveAddins Array("Web")
	qtApp.Launch ' Start QuickTest
	qtApp.Visible = False ' Make the QuickTest application visible
	' Set QuickTest run options
	qtApp.Options.Run.RunMode = "Fast"
	qtApp.Options.Run.ViewResults = False
	qtApp.Open Test_path, True ' Open the test in read-only mode
	' set run settings for the test
	Set qtTest = qtApp.Test
	qtTest.Run ' Run the test
	qtTest.Close ' Close the test
	qtApp.quit
	Set qtTest = Nothing ' Release the Test object
	Set qtApp = Nothing ' Release the Application object
	MsgBox "Defect(s) logged into Jira!"
Else
	MsgBox "Execution done sucessuflly, no bug found"
End If
