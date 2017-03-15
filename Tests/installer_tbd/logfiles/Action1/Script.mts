'##########################################################################################################
' Objective: Verify that  all log files of trust are getting created 
' Test Description: All folders and log files should get created  under system temp \Mobile Labs folder.
' Steps:
' Step1: Locate logfiles and folders under temp\MobileLabs
'##########################################################################################################

'#######################################################
'Declare Variables
Dim	intFileNotFoundIndex
Dim intUsedRowsCount
Dim intWorkSheetsCounter
Dim strLogFilesExcelPath
Dim strStepsToReproduce
Dim strStepName
Dim strSummary
Dim strActualResult
Dim intStep
Dim booleanAllFilesFoldersFound
Dim arrStepName()
Dim arrExpectedResult()
Dim arrActualResult()
Dim arrDetail()
Dim arrStatus()
Dim arrFilesNotFound()

Environment("FileFound") = False
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

'Launch Object Manager to create corresponding logs
If LaunchTrustOM = True Then
	If Window("regexpwndtitle:=Mobile Labs Trust").Exist(2) Then 
		Window("regexpwndtitle:=Mobile Labs Trust").Close
	ElseIf WpfWindow("regexpwndtitle:=Mobile Labs Trust", "devname:=Root").Exist(2) Then
		WpfWindow("regexpwndtitle:=Mobile Labs Trust", "devname:=Root").Close
	End If
End If

'Calling function to get the path of Trust
StrTrustPath = GetTrustInstallDir

'Launch Device Viewer to create corresponding logs
Set objWScript = CreateObject("WScript.Shell")
objWScript.Run  "cmd.exe /C " & Left(StrTrustPath,2) & "& cd " & StrTrustPath & "& DeviceViewer.exe"

If Window("regexpwndtitle:=Device Viewer").Exist Then
	Window("regexpwndtitle:=Device Viewer").Close
End If 

'#######################################################

intStep = intStep+1
strStepName = "Go to the system temp folder and locate all log files"
'*********************************************************************************************************************
' Step1: Locate Trust Log Files
'All Trust log files should be created in system temp folder

intSubStep =1

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Locate Trust Log Files."

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "All Trust Log Files should be present in the system temp folder."

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentPath = Environment("TestDir")

booleanAllFilesFoldersFound = True

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

'Pointing to the excel consisting of all log files details that should be present in temp folder
Set objExcel = objExcel.Workbooks.Open(GetFilePath("InstallerDirs.xlsx"))
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objWorkSheet = objExcel.Worksheets("Temp_MobileLabs")

'Setting intial value of steps to reproduce
strStepsToReproduce = strStepsToReproduce & strStepName & "."& VBNewLine


For intUsedRowsCount= 2 To objWorkSheet.UsedRange.Rows.Count

	strFolderOrFilePath = objWorkSheet.Cells(intUsedRowsCount,1)
	tempArr =  Split (objWorkSheet.Cells(intUsedRowsCount,1).Value , "\")

	If  objWorkSheet.Cells(intUsedRowsCount,2) = "File" Then
		strStepsToReproduce = strStepsToReproduce & intSubStep & ": " & "Locate log file " & tempArr(ubound(tempArr)) & " under folder: " & tempArr(ubound(tempArr)-1)&"."& VBNewLine
		If  Not (objFSO.FileExists(GetSystemDrivePath(strFolderOrFilePath))) Then
			'Setting the flag to False if  any one of the log files is not located
			booleanAllFilesFoldersFound = False

			'Forming an array of the log files not located
			ReDim Preserve arrFilesNotFound(intFileNotFoundIndex)
			arrFilesNotFound(intFileNotFoundIndex)= tempArr(ubound(tempArr))
			intFileNotFoundIndex = intFileNotFoundIndex + 1
		End If
	Else 
		If  objWorkSheet.Cells(intUsedRowsCount,2) = "Folder" Then
			strStepsToReproduce = strStepsToReproduce & intSubStep & ": " & "Locate folder " & tempArr(ubound(tempArr)) & " under folder: " & tempArr(ubound(tempArr)-1)&"."& VBNewLine
			If  Not (objFSO.FolderExists(GetSystemDrivePath(strFolderOrFilePath))) Then
				'Setting the flag to False if  any one of the log files is not located
				booleanAllFilesFoldersFound = False
				
				'Forming an array of the log files not located
				ReDim Preserve arrFilesNotFound(intFileNotFoundIndex)
				arrFilesNotFound(intFileNotFoundIndex) = tempArr(ubound(tempArr))
				intFileNotFoundIndex = intFileNotFoundIndex + 1
			End If
		End If
	End If
	intSubStep = intSubStep +1 
Next

'Destroying the objects 
objExcel.Close
Set objExcel = Nothing
Set objFSO = Nothing
Set objWScript = Nothing 
Set objWorkSheet = Nothing 


If  booleanAllFilesFoldersFound  = True Then

	ReDim Preserve arrActualResult(intStep)
	strActualResult = "All log files and folders of Trust located in system temp folder."
	arrActualResult(intStep) = strActualResult
	strSummary = "All log files and folders of Trust located."&VBNewLine 
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

Else

	strActualResult = ""
	For intFileNotFoundIndex = 0 to ubound(arrFilesNotFound)
		strActualResult  = strActualResult &intFileNotFoundIndex+1& ": "&arrFilesNotFound(intFileNotFoundIndex)&VBNewLine
	Next
	ReDim Preserve arrActualResult(intStep)
	strActualResult = "The following trust log files/folders are not located in system temp folder: "&VBNewLine&strActualResult
	arrActualResult(intStep) = strActualResult
	ReDim Preserve arrDetail(intStep)
	strSummary = "All log files and folders of Trust are not located."&VBNewLine 
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	arrStatus(intStep) = "Failed"

End If


''*********************************************************************************************************************

'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus
