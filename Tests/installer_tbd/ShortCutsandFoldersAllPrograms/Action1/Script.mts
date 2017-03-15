'##########################################################################################################
' Objective: Verify that  all Trust  Files ShortCut and Folder are Present  System All Programs .
' Test Description: Shortcuts for All Files should be Present under System All Programs .
' Steps:

' Step1: Verify Files Shotcut & Folder (Trust)

'##########################################################################################################
 
'#######################################################
'Declare Variables
Dim intFileNotFoundIndex
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

'#######################################################
 
intStep = intStep+1
strStepName = "Go to the system All Program and locate all Shotcut Files and Folders"
'*********************************************************************************************************************
' Step1: Locate Trust ShortCut Files
'Trust ShortCut Files should be present @ %\ProgramData\Microsoft\Windows\Start Menu\Programs\Mobile Labs
 
intSubStep =1
 
ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Locate Trust ShortCut Files & Folder."
 
ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "All Trust Folder should be present in the System Start menu-->All Program."
 
ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
'strRootFolder = ""
strCurrentPath = Environment("TestDir")
'strInstallerPath = ""
 
'Loop until "MobileLabs Automation Framework" folder is found
'blnParentFolderFound = True
'Do While Replace(Split(strCurrentPath,"\")(UBound(Split(strCurrentPath,"\"))), " ", "") <> "MobileLabsAutomationFramework"
'                strCurrentPath = objFSO.GetParentFolderName(strCurrentPath)
'                'Exit if reaches the system drive
'                If InStr(1, strCurrentPath, "\") = 0 Then
'                                blnParentFolderFound = False
'                                Exit Do
'                End If
'Loop
' 
''Define the path of the Root Folder: <MobileLabs Automation Framework>
'If blnParentFolderFound Then
'                If Right(strCurrentPath,1) <> "\" Then
'                                strRootFolder = strCurrentPath & "\"
'                End If
'Else
'                ExitTest
'End If
 
booleanAllFilesFoldersFound = True
 
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
 
'Pointing to the excel consisting of all log files details that should be present in temp folder
Set objExcel = objExcel.Workbooks.Open(GetFilePath("InstallerDirs.xlsx"))
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
'Pointing to the LogFile worksheet
Set objWorkSheet = objExcel.Worksheets("Programs_MobileLabs")
 
'Setting intial value of steps to reproduce
strStepsToReproduce = strStepsToReproduce & strStepName & "."& VBNewLine
 
 
For intUsedRowsCount= 2 To objWorkSheet.UsedRange.Rows.Count
 
                strFolderOrFilePath = objWorkSheet.Cells(intUsedRowsCount,1)
                tempArr =  Split (objWorkSheet.Cells(intUsedRowsCount,1).Value , "\")
 
                If  objWorkSheet.Cells(intUsedRowsCount,2) = "File" Then
                                strStepsToReproduce = strStepsToReproduce & intSubStep & ": " & "Locate File ShortCut " & tempArr(ubound(tempArr)) & " under folder: " & tempArr(ubound(tempArr)-1)&"."& VBNewLine
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
Set  objExcel = Nothing
Set objFSO = Nothing
Set objWScript = Nothing 
Set objWorkSheet = Nothing 
 
 
If  booleanAllFilesFoldersFound  = True Then
 
                ReDim Preserve arrActualResult(intStep)
                strActualResult = "All Trust ShortCut  Files & folders are Located @ System -->StartMenu-->All Programs ."
                arrActualResult(intStep) = strActualResult
                strSummary = "All Trust ShortCut  Files & folders are Located."&VBNewLine 
                ReDim Preserve arrDetail(intStep)
                arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
                "Steps to Reproduce:" & VBNewLine & strStepsToReproduce
 
Else
 
                strActualResult = ""
                For intFileNotFoundIndex = 0 to ubound(arrFilesNotFound)
                                strActualResult  = strActualResult &intFileNotFoundIndex+1& ": "&arrFilesNotFound(intFileNotFoundIndex)&VBNewLine
                Next
                ReDim Preserve arrActualResult(intStep)
                strActualResult = "The following Trust ShortCut  Files & folders are not located in System Start menu-->All Program: "&VBNewLine&strActualResult
                arrActualResult(intStep) = strActualResult
                ReDim Preserve arrDetail(intStep)
                strSummary = "Trust ShortCut  Files & folders not located."&VBNewLine 
                arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
                "Steps to Reproduce:" & VBNewLine & strStepsToReproduce
                arrStatus(intStep) = "Failed"
 
End If
 
 
''*********************************************************************************************************************
 
'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus