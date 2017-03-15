'##########################################################################################################
' Objective: Installer should launch with correct UI and remove trust without any errors
' Test Description: 
' Steps:
' Step1: Launch the setup through setup.exe.
' Step2: Hit "Next" button on the first screen.
' Step3: Click the 'Remove' button.
' Step4: Click the 'Remove' button on Ready to remove mobile labs trust screen.
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
Dim booleanProceedToNextStep
Dim arrStepName()
Dim arrExpectedResult()
Dim arrActualResult()
Dim arrDetail()
Dim arrStatus()
'#######################################################

'#######################################################
'Initializations
intStep = 0
booleanProceedToNextStep = True

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

''Install Trust if  not already installed
InstallTrust

'Locate the setup.exe file
Set objWScript = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strRootFolder = ""
strCurrentPath = Environment("TestDir")
strInstallerPath = ""

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
		strRootFolder = strCurrentPath & "\"
	End If
Else
	ExitTest
End If

'Loop until Installation-Media folder is not found
strInstallerPath = strRootFolder & "TrustBuilds"
blnFileFound = False
Set objFolder = objFSO.GetFolder(strInstallerPath).SubFolders
Do While objFolder.Count > 0 AND Not(blnFileFound)
	For Each objSubFolder in objFolder
		Do While Not(blnFileFound)
			For Each objChildFolder in objSubFolder.SubFolders
				If InStr(1, objChildFolder.Name, "Installation", 1) > 0 Then
					strInstallerPath = objFSO.GetFolder(objChildFolder.Path) & "\setup.exe"
					blnFileFound = True
					Exit For
				End If
			Next
		Loop
	Next
Loop

'#######################################################

intStep = intStep+1
'*********************************************************************************************************************
' Step1: Launch the setup through setup.exe.
'Expected Result: Welcome screen should launch with Back button disabled and Next and Cancel buttons enabled.

ReDim Preserve arrStepName(intStep)
arrStepName(intStep) = "Launch the setup through setup.exe."
strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Launch the setup through setup.exe." & VBNewLine

ReDim Preserve arrExpectedResult(intStep)
arrExpectedResult(intStep) = "Welcome screen should launch with Back button disabled and Next and Cancel buttons enabled."

ReDim Preserve arrStatus(intStep)
arrStatus(intStep) = "Passed"

'Launch installer
SystemUtil.Run strInstallerPath

'Check if the installer window launches
If Window("Mobile Labs Trust Setup").Exist(2) Then
	strSummary = "Installer window launched successfully."
	strActualResult = "Welcome screen is being displayed correctly."
	blnResult = True

	If Not(Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled")) Then
		strSummary = strSummary & " Back button is disabled."
		strActualResult = strActualResult & " Back button is disabled."
	Else
		blnResult = False
		strSummary = strSummary & " Back button is enabled."
		strActualResult = strActualResult & " Back button is enabled."
	End If

	If Window("Mobile Labs Trust Setup").WinButton("Next").GetROProperty("enabled") Then
		strSummary = strSummary & " Next button is enabled."
		strActualResult = strActualResult & " Next button is enabled."
	Else
		blnResult = False
		booleanProceedToNextStep = False 
		strSummary = strSummary & " Next button is disabled."
		strActualResult = strActualResult & " Next button is disabled."
	End If

	If Window("Mobile Labs Trust Setup").WinButton("Cancel").GetROProperty("enabled") Then
		strSummary = strSummary & " Cancel button is enabled."
		strActualResult = strActualResult & " Cancel button is enabled."
	Else
		blnResult = False
		strSummary = strSummary & " Cancel button is disabled."
		strActualResult = strActualResult & " Cancel button is disabled."
	End If

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = strActualResult
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary

	If blnResult Then
		Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
	Else
		Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
	End If
	
Else
	strSummary = "Installer window did not launch."
	arrStatus(intStep) = "Failed"

	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "Welcome screen is not being displayed."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
End If


intStep = intStep+1
'*********************************************************************************************************************
'Step2: Click  the  'Next' button.
'Expected Result: . Destination Change,Repair or Remove installation screen should come up, with Change , and Next buttons disabled , and Repair ,Remove ,Back  and Cancel buttons enabled 
If booleanProceedToNextStep Then

	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Click 'Next' button."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Click 'Next' button on the welcome screen." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "Destination screen : 'Change , repair or remove installation' should launch with Repair , Remove , Back and Cancel buttons enabled , and Next and Change buttons disabled."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	blnResult = True
	
	'Click on Next button on the Welcome screen , to go to next screen
	
	Window("Mobile Labs Trust Setup").WinButton("Next").Click

	'Check if the next screen has been brought up
	If  Window("Mobile Labs Trust Setup").Static("ChangeRepairORRemove").GetROProperty("text") = "Change, repair, or remove installation" Then
		strSummary = "'Change , repair or remove installation' screen is being displayed."
		strActualResult = "'Change , repair or remove installation'  screen has been brought up."

		'Check whether Repair , Remove , Back and Cancel buttons are enabled , and Next and Change buttons are disabled."
		If  Window("Mobile Labs Trust Setup").WinButton("Repair").GetROProperty("enabled") Then
			strSummary = strSummary & " Repair  button is enabled."
			strActualResult = strActualResult & " Repair button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Repair  button is disabled."
			strActualResult = strActualResult & " Repair button is disabled."
		End If

		If  Window("Mobile Labs Trust Setup").WinButton("Remove").GetROProperty("enabled") Then
			strSummary = strSummary & " Remove  button is enabled."
			strActualResult = strActualResult & " Remove button is enabled."
		Else
			blnResult = False
			booleanProceedToNextStep = False
			strSummary = strSummary & " Remove  button is disabled."
			strActualResult = strActualResult & " Remove button is disabled."
		End If

		If  Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled") Then
			strSummary = strSummary & " Back  button is enabled."
			strActualResult = strActualResult & " Back button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Back  button is disabled."
			strActualResult = strActualResult & " Back button is disabled."
		End If

		If  Window("Mobile Labs Trust Setup").WinButton("Cancel").GetROProperty("enabled") Then
			strSummary = strSummary & " Cancel  button is enabled."
			strActualResult = strActualResult & " Cancel button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Cancel  button is disabled."
			strActualResult = strActualResult & " Cancel button is disabled."
		End If

		If Not Window("Mobile Labs Trust Setup").WinButton("Change").GetROProperty("enabled") Then
			strSummary = strSummary & " Change  button is disabled."
			strActualResult = strActualResult & " Change button is disabled."
		Else
			blnResult = False
			strSummary = strSummary & " Change  button is enabled."
			strActualResult = strActualResult & " Change button is enabled."
		End If

		If Not Window("Mobile Labs Trust Setup").WinButton("Next").GetROProperty("enabled") Then
			strSummary = strSummary & " Next  button is disabled."
			strActualResult = strActualResult & " Next button is disabled."
		Else
			blnResult = False
			strSummary = strSummary & " Next  button is enabled."
			strActualResult = strActualResult & " Next button is enabled."
		End If

		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = strActualResult
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary

		If blnResult Then
			Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
		Else
			Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
		End If

	Else
	
	arrStatus(intStep) = "Failed"
	strSummary = " 'Change , repair or remove installation' screen is not being displayed."
	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = "'Change , repair or remove installation'  screen has not  been brought up."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
		
		
	End If


End If


intStep = intStep+1
'*********************************************************************************************************************
'Step3: Click  the  'Remove' button.
'Expected Result: . Destination 'Ready to remove Mobile Labs Trust ' screen should come up, with Back , Remove and Cancel buttons enabled 
If booleanProceedToNextStep Then

	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Click 'Remove' button on'Change , repair or remove installation' screen."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Click 'Remove' button on the 'Change , repair or remove installation'  screen." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "Destination screen :Ready to remove Mobile Labs Trust ' should launch with Back , Remove and Cancel buttons enabled."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	blnResult = True
	
	'Click on Remove button on the 'Change , repair or remove installation' screen , to go to next screen
	
	Window("Mobile Labs Trust Setup").WinButton("Remove").Click

	'Check if the next screen has been brought up
	If  Window("Mobile Labs Trust Setup").Static("ReadyToRemove").GetROProperty("text") = "Ready to remove Mobile Labs Trust" Then
		strSummary = " 'Ready to remove Mobile Labs Trust' screen is being displayed."
		strActualResult = "'Ready to remove Mobile Labs Trust'  screen has been brought up."

		'Check whether  Back , Remove and Cancel buttons are enabled 
		If  Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled") Then
			strSummary = strSummary & " Back  button is enabled."
			strActualResult = strActualResult & " Back button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Back  button is disabled."
			strActualResult = strActualResult & " Back button is disabled."
		End If

		If  Window("Mobile Labs Trust Setup").WinButton("Remove").GetROProperty("enabled") Then
			strSummary = strSummary & " Remove  button is enabled."
			strActualResult = strActualResult & " Remove button is enabled."
		Else
			blnResult = False
			booleanProceedToNextStep = False
			strSummary = strSummary & " Remove  button is disabled."
			strActualResult = strActualResult & " Remove button is disabled."
		End If

		If  Window("Mobile Labs Trust Setup").WinButton("Cancel").GetROProperty("enabled") Then
			strSummary = strSummary & " Cancel  button is enabled."
			strActualResult = strActualResult & " Cancel button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Cancel  button is disabled."
			strActualResult = strActualResult & " Cancel button is disabled."
		End If

		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = strActualResult
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary

		If blnResult Then
			Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
		Else
			Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
		End If

	Else
	
	arrStatus(intStep) = "Failed"
	strSummary = " 'Ready to remove Mobile Labs Trust' screen is not being displayed."
	ReDim Preserve arrActualResult(intStep)
	arrActualResult(intStep) = " 'Ready to remove Mobile Labs Trust' screen has not  been brought up."
	
	ReDim Preserve arrDetail(intStep)
	arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
	"Steps to Reproduce:" & VBNewLine & strStepsToReproduce

	Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
		
		
	End If


End If



intStep = intStep+1
'*********************************************************************************************************************
'Step4: Click the 'Remove' button on Ready to remove mobile labs trust screen.
'Expected Result: . : Removing Mobile Labs Trust message should be displayed until uninstallation completes. 
'Once complete the 'Completed the Mobile Labs Trust Setup Wizard'  screen should appear 
' with Finish button enabled and Back and Cancel buttons disabled.

If booleanProceedToNextStep Then

	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Click 'Remove' button on 'Ready to remove Mobile Labs Trust' screen ."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Click 'Remove' button on the 'Ready to remove Mobile Labs Trust' screen." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "Removing Mobile Labs Trust message should be displayed until uninstallation completes." &_
	"Once complete the 'Completed the Mobile Labs Trust Setup Wizard'  screen should appear with Finish button enabled and Back and Cancel buttons disabled."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	blnResult = True
	
	'Click on Remove button on the 'Ready to remove mobile labs trust screen' screen , to go to next screen
	
	Window("Mobile Labs Trust Setup").WinButton("Remove").Click

	'Check if the  'Removing Mobile Labs Trust ' screen has been brought up

	If Window("Mobile Labs Trust Setup").Static("Removing").GetROProperty("text") = "Removing Mobile Labs Trust" Then

		Do While Window("Mobile Labs Trust Setup").Static("Removing").Exist(1)
			' Wait until uninstallation completes
		Loop

		If Window("Mobile Labs Trust Setup").Static("Completed").GetROProperty("text") = "Completed the Mobile Labs Trust Setup Wizard" Then

			strSummary = "Uninstallation completion screen is being displayed."
			strActualResult = "Uninstallation completion screen has been brought up."
			blnResult = True
		
			If Not(Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled")) Then
				strSummary = strSummary & " Back button is disabled."
				strActualResult = strActualResult & " Back button is disabled."
			Else
				blnResult = False
				strSummary = strSummary & " Back button is enabled."
				strActualResult = strActualResult & " Back button is enabled."
			End If

			If Not(Window("Mobile Labs Trust Setup").WinButton("Cancel").GetROProperty("enabled")) Then
				strSummary = strSummary & " Cancel button is disabled."
				strActualResult = strActualResult & " Cancel button is disabled."
			Else
				blnResult = False
				strSummary = strSummary & " Cancel button is enabled."
				strActualResult = strActualResult & " Cancel button is enabled."
			End If

			If Window("Mobile Labs Trust Setup").WinButton("Finish").GetROProperty("enabled") Then
				strSummary = strSummary & " Finish button is enabled."
				strActualResult = strActualResult & " Finish button is enabled."
			Else
				blnResult = False
				booleanProceedToNextStep = False
				strSummary = strSummary & " Back button is disabled."
				strActualResult = strActualResult & " Back button is disabled."
			End If

			ReDim Preserve arrActualResult(intStep)
			arrActualResult(intStep) = strActualResult
			ReDim Preserve arrDetail(intStep)
			arrDetail(intStep) = "Summary: " & strSummary
		
			If blnResult Then
				Reporter.ReportEvent micPass, arrStepName(intStep), arrActualResult(intStep)
			Else
				Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
			End If
	
		Else
		
			arrStatus(intStep) = "Failed"
			strSummary = " 'Completed the Mobile Labs Trust Setup Wizard' screen is not being displayed."
			ReDim Preserve arrActualResult(intStep)
			arrActualResult(intStep) = " 'Completed the Mobile Labs Trust Setup Wizard' screen has not  been brought up."
			
			ReDim Preserve arrDetail(intStep)
			arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
			"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
		
			Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
		End If	
	Else
		arrStatus(intStep) = "Failed"
		strSummary = " 'Completed the Mobile Labs Trust Setup Wizard' screen is not being displayed."
		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = " 'Removing Mobile Labs Trust' screen has not  been brought up."
		
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
		"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	
		Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)		
	End If

End If

'*********************************************************************************************************************

'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus

''Close the installation wizard
If booleanProceedToNextStep = True Then
	Window("Mobile Labs Trust Setup").WinButton("Finish").Click
End If

Set objWScript = Nothing
Set objFSO = Nothing
