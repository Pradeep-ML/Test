'##########################################################################################################
' Objective: Verify that  all files and subfolders created on installation
' Test Description: All files and subfolders created under Program Files\MobileLabs
' Steps:
' Step1: Locate all files and folders under Program Files\MobileLabs
'##########################################################################################################

'#######################################################
'Declare Variables
Dim strNotFound
Dim intUsedRowsCount 
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
Dim i

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

'Get Trust installation directory and determine the Program Files folder name
strProgramFiles = ""
arrPath = Split(GetTrustInstallDir, "\")
For iCount = 0 To UBound(arrPath)
	If InStr(1, arrPath(iCount), "Program Files", 1) > 0 Then
		strProgramFiles = Replace(arrPath(iCount), " ", "")
		Exit For
	End If
Next

'#######################################################

intStep = intStep+1
strStepName = "Go to the Program Files and locate MobileLabs subfolders and files"
'*********************************************************************************************************************
' Step1: Locate Trust files and sub folders
'All Trust files and folders should be created under Program Files\MobileLabs

intSubStep =1

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Locate Trust Log Files."

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "All Trust Log Files should be present in the system temp folder."

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"


Set objFSO = CreateObject("Scripting.FileSystemObject")

strCurrentPath = Environment("TestDir")

boolAllFilesFoldersFound = True 

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False

'Open InstallerDirs.xlsx containing files and folders path

Set objExcel = objExcel.Workbooks.Open(GetFilePath("InstallerDirs.xlsx"))
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Open Programfiles_MobileLabs sheet
Set objWorkSheet = objExcel.Worksheets(strProgramFiles & "_MobileLabs")

'Setting intial value of steps to reproduce
strStepsToReproduce = strStepsToReproduce & strStepName & "."& VBNewLine

intUsedRowsCount = objWorkSheet.UsedRange.Rows.Count
For i = 2 To intUsedRowsCount 
 
 strFileorFolderPath = objWorkSheet.Cells(i,1)
 progArr =  Split (objWorkSheet.Cells(i,1).Value , "\")
	If  LCase(objWorkSheet.Cells(i,2).Value) = "file" Then
		If Not (objFSO.FileExists(GetSystemDrivePath(strFileorFolderPath))) Then
		'If Not (objFSO.FileExists(strFileorFolderPath)) Then
			boolAllFilesFoldersFound = False
			'Forming an array of the log files not located
			ReDim Preserve arrFilesNotFound(intFileNotFoundIndex)
			arrFilesNotFound(intFileNotFoundIndex)= progArr(ubound(progArr)) 
			  intFileNotFoundIndex = intFileNotFoundIndex + 1
		Else
			strStepsToReproduce = strStepsToReproduce & intSubStep & ": " & "Verified file " & progArr(ubound(progArr)) & " under folder: " & progArr(ubound(progArr)-1)&"."& VBNewLine
		End If

	ElseIf   LCase(objWorkSheet.Cells(i,2).Value)  = "folder" Then
	
		If Not (objFSO.FolderExists(GetSystemDrivePath(strFileorFolderPath))) Then
		'If Not (objFSO.FileExists(strFileorFolderPath)) Then
			boolAllFilesFoldersFound = False
			'Forming an array of the log files not located
			ReDim Preserve arrFilesNotFound(intFileNotFoundIndex)
			 arrFilesNotFound(intFileNotFoundIndex)= progArr(ubound(progArr)) 
			 intFileNotFoundIndex = intFileNotFoundIndex + 1
		Else
			strStepsToReproduce = strStepsToReproduce & intSubStep & ": " & "Verified sub folder " & progArr(ubound(progArr)) & " under folder: " & progArr(ubound(progArr)-1)&"."& VBNewLine
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


If  boolAllFilesFoldersFound  = True Then

	 ReDim Preserve arrActualResult(intStep)
	 strActualResult = "All files and sub folders of Trust located under Program Files\MobileLabs."
	 arrActualResult(intStep) = strActualResult
	 strSummary = "All files and sub folders of Trust located."&VBNewLine 
	 ReDim Preserve arrDetail(intStep)
	 arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	 "Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	Else
	
	 strActualResult = ""
	 For intFileNotFoundIndex = 0 to ubound(arrFilesNotFound)
	  strActualResult  = strActualResult &intFileNotFoundIndex+1& ": "&arrFilesNotFound(intFileNotFoundIndex)&VBNewLine
	 Next
	 ReDim Preserve arrActualResult(intStep)
	 strActualResult = "The following trust files and folders not found at location Program Files\MobileLabs: "&VBNewLine&strActualResult
	 arrActualResult(intStep) = strActualResult
	 ReDim Preserve arrDetail(intStep)
	 strSummary = "All l files and sub folders of Trust not located.."&VBNewLine 
	 arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	 "Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	 arrStatus(intStep) = "Failed"

End If


''*********************************************************************************************************************

'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus




