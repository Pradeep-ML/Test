'##########################################################################################################
' Objective: Verify that shortcuts for Object Manager and Device Viewer are placed on desktop
' Test Description: Shortcuts for Object Manager and Device Viewer should be placed on the public or all users desktop
' Steps:
' Step1: Locate Object Manager shortcut.
' Step2: Locate Device Viewer shortcut.
'##########################################################################################################

'#######################################################
'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim strSummary
Dim strActualResult
Dim intStep
Dim blnResult
Dim strTestName
Dim arrStepName()
Dim arrExpectedResult()
Dim arrActualResult()
Dim arrDetail()
Dim arrStatus()
'#######################################################

'#######################################################
'Initializations
intStep = 0

ReDim Preserve arrStepName(intStep)
arrStepName(0) = "Step Name"

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(0) = "Expected Result"

ReDim Preserve arrActualResult(intStep)
arrActualResult(0) = "Actual Result"

ReDim Preserve arrDetail(intStep)
arrDetail(0) = "Details"

ReDim Preserve arrStatus(intStep)
arrStatus(0) = "Status"
'#######################################################

'#######################################################
'Initial Setup

'Install Trust if not already installed
InstallTrust

'#######################################################

intStep = intStep+1
'*********************************************************************************************************************
' Step1: Locate Object Manager shortcut.
'Expected Result: Object Manager shortcut should be there on the Public Desktop.

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Locate Object Manager shortcut."
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Locate Object Manager shortcut." & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "Object Manager shortcut should be there on the Public Desktop."

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

'Verify if Locate Object Manager shortcut is there on the Public Desktop
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFiles = objFSO.GetFolder(GetPublicDesktopPath).Files
blnFound = False
For Each objFile in objFiles
	If Replace(UCase(objFile.Name), " ", "") = "OBJECTMANAGER.LNK" Then
		blnFound = True
	End If
Next

'Report results
If blnFound Then
	strSummary = "Object Manager shortcut found."
	strActualResult = "Object Manager shortcut is located on the Public Desktop."

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary

	Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
	
Else
	strSummary = "Object Manager shortcut was not found."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Object Manager shortcut doesn't exist on the Public Desktop located at - " & GetPublicDesktopPath
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If

'*********************************************************************************************************************

intStep = intStep+1
'*********************************************************************************************************************
' Step2: Locate Device Viewer shortcut.
'Expected Result: Device Viewer shortcut should be there on the Public Desktop.

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Locate Device Viewer shortcut."
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Locate Device Viewer shortcut." & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "Device Viewer shortcut should be there on the Public Desktop."

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

'Verify if Locate Object Manager shortcut is there on the Public Desktop
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFiles = objFSO.GetFolder(GetPublicDesktopPath).Files
blnFound = False
For Each objFile in objFiles
	If Replace(UCase(objFile.Name), " ", "") = "DEVICEVIEWER.LNK" Then
		blnFound = True
	End If
Next

'Report results
If blnFound Then
	strSummary = "Device Viewer shortcut found."
	strActualResult = "Device Viewer shortcut is located on the Public Desktop."

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary

	Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
	
Else
	strSummary = "Device Viewer shortcut was not found."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Device Viewer shortcut doesn't exist on the Public Desktop located at - " & GetPublicDesktopPath
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If

'*********************************************************************************************************************

'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus

'Destroy objects
Set objFSO = Nothing
Set objFiles = Nothing