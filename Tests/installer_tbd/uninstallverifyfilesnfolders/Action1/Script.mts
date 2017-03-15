'##########################################################################################################
' Objective: All files and folders should be removed after Trust is uninstalled
' Test Description: Uninstall Trust and verify that all files and folders are removed
' Steps:
' Step1: Verify 'Mobile Labs' folder is removed from the All Programs
' Step2: Verify 'Mobile Labs' folder is removed from the Program Files
' Step3: Verify 'Mobile Labs' folder is removed from %AppData/Roaming
' Step4: Verify 'Mobile Labs' folder is removed from %AppData\Local\Temp
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

'Uninstall Trust if already installed
UninstallTrust

Set objFSO = CreateObject("Scripting.FileSystemObject")
'#######################################################

intStep = intStep+1
'*********************************************************************************************************************
' Step1: Verify 'Mobile Labs' folder is removed from the All Programs
'Expected Result: 'Mobile Labs' folder shouldn't exist in Windows All Programs directory

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Verify 'Mobile Labs' folder is removed from the All Programs"
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Locate 'Mobile Labs' in Windows All Programs" & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "'Mobile Labs' folder shouldn't exist in Windows All Programs directory"

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

'Check if Mobile Labs folder exists in the All Programs directory
If Not(objFSO.FolderExists(GetAllProgramsFolderPath & "Mobile Labs")) Then
	strSummary = "'Mobile Labs' folder has been successfully removed from the All Programs directory."
	strActualResult = "Uninstalling Trust has removed the 'Mobile Labs' folder from the Windows All Programs directory."

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult

	Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
Else
	strSummary = "'Mobile Labs' folder has not been removed from the All Programs directory."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Uninstalling Trust has not removed the 'Mobile Labs' folder from the Windows All Programs directory."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If

'*********************************************************************************************************************

intStep = intStep+1
'*********************************************************************************************************************
' Step2: Verify 'Mobile Labs' folder is removed from the Program Files
'Expected Result: 'Mobile Labs' folder shouldn't exist in Windows Program Files directory

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Verify 'Mobile Labs' folder is removed from the Program Files"
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Locate 'Mobile Labs' in Windows Program Files" & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "'Mobile Labs' folder shouldn't exist in Windows Program Files directory"

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

arrFolderPath = Split(GetTrustInstallDir, "\")
strFolderPath = ""
For i = 0 To UBound(arrFolderPath)
	If LCase(Replace(arrFolderPath(i), " ", "")) = "trust" Then
		Exit For
	End If
	strFolderPath = strFolderPath & arrFolderPath(i) & "\"
Next

'Check if Mobile Labs folder exists in the Program Files directory
If Not(objFSO.FolderExists(strFolderPath)) Then
	strSummary = "'Mobile Labs' folder has been successfully removed from the Program Files directory."
	strActualResult = "Uninstalling Trust has removed the 'Mobile Labs' folder from the Windows Program Files directory."

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult

	Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
Else
	strSummary = "'Mobile Labs' folder has not been removed from the Program Files directory."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Uninstalling Trust has not removed the 'Mobile Labs' folder from the Windows Program Files directory."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If

'*********************************************************************************************************************

intStep = intStep+1
'*********************************************************************************************************************
' Step3: Verify 'Mobile Labs' folder is not removed from %AppData/Roaming
'Expected Result: 'Mobile Labs' folder shouldn't be removed from Windows %AppData/Roaming directory

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Verify 'Mobile Labs' folder is not removed from the %AppData/Roaming"
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Locate 'Mobile Labs' in Windows %AppData/Roaming" & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "Mobile Labs' folder shouldn't be removed from Windows %AppData/Roaming directory"

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

arrFolderPath = Split(Environment("SystemTempDir"), "\")
strFolderPath = ""
For i = 0 To UBound(arrFolderPath)
	If LCase(Replace(arrFolderPath(i), " ", "")) = "local" Then
		Exit For
	End If
	strFolderPath = strFolderPath & arrFolderPath(i) & "\"
Next

strFolderPath = strFolderPath & "Roaming"

'Check if Mobile Labs folder exists in the Program Files directory
If objFSO.FolderExists(strFolderPath) Then
	strSummary = "'Mobile Labs' folder is therein the %AppData/Roaming directory."
	strActualResult = "Uninstalling Trust has not removed the 'Mobile Labs' folder from the Windows %AppData/Roaming directory."

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult

	Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
Else
	strSummary = "'Mobile Labs' folder has been removed from the %AppData/Roaming directory."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Uninstalling Trust has removed the 'Mobile Labs' folder from the Windows %AppData/Roaming directory."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If

'*********************************************************************************************************************

intStep = intStep+1
'*********************************************************************************************************************
' Step4: Verify 'Mobile Labs' folder is removed from %AppData\Local\Temp
'Expected Result: Mobile Labs' folder shouldn't exist in Windows %AppData\Local\Temp directory

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Verify 'Mobile Labs' folder is removed from the %AppData\Local\Temp"
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Locate 'Mobile Labs' in Windows %AppData\Local\Temp" & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "'Mobile Labs' folder shouldn't exist in Windows %AppData\Local\Temp directory"

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

strFolderPath = Environment("SystemTempDir") & "\Mobile Labs"

'Check if Mobile Labs folder exists in the Program Files directory
If Not(objFSO.FolderExists(strFolderPath)) Then
	strSummary = "'Mobile Labs' folder has been successfully removed from the %AppData\Local\Temp directory."
	strActualResult = "Uninstalling Trust has removed the 'Mobile Labs' folder from the Windows %AppData\Local\Temp directory."

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult

	Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
Else
	strSummary = "'Mobile Labs' folder has not been removed from the %AppData\Local\Temp directory."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Uninstalling Trust has not removed the 'Mobile Labs' folder from the Windows %AppData\Local\Temp directory."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If

'*********************************************************************************************************************

'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus

'Destroy the objects
Set objFSO = Nothing