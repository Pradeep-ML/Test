'##########################################################################################################
' Objective: Installer should launch with correct UI and install without any errors
' Test Description: Installation should commit without any errors and all windows should open with correct UI
' Steps:
' Step1: Launch the setup through setup.exe.
' Step2: Hit "Next" button on the first screen.
' Step3: Accept the License and hit "Next" button.
' Step4: Keep the default installation location and hit "Next" button.
' Step5: Hit "Install" button.
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

'*********************************************************************************************************************
'Proceed only if the Installer launches successfully
If blnResult Then

	intStep = intStep+1
	'*********************************************************************************************************************
	' Step2: Hit "Next" button on the first screen.
	'Expected Result: EULA window should launch with Print, Back and Cancel buttons enabled and Next button disabled.
	
	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Hit 'Next' button on the first screen."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Hit 'Next' button on the first screen." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "EULA window should launch with Print, Back and Cancel buttons enabled and Next button disabled."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	'Hit Next button to go to the next screen
	Window("Mobile Labs Trust Setup").WinButton("Next").Click
	
	'Check if the EULA is brought up
	If Window("Mobile Labs Trust Setup").Static("EULA").GetROProperty("text") = "End-User License Agreement" Then
		strSummary = "EULA screen is being displayed."
		strActualResult = "EULA screen has been brought up."
		blnResult = True
	
		If Window("Mobile Labs Trust Setup").WinButton("Print").GetROProperty("enabled") Then
			strSummary = strSummary & " Print button is enabled."
			strActualResult = strActualResult & " Print button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Print button is disabled."
			strActualResult = strActualResult & " Print button is disabled."
		End If
	
		If Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled") Then
			strSummary = strSummary & " Back button is enabled."
			strActualResult = strActualResult & " Back button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Back button is disabled."
			strActualResult = strActualResult & " Back button is disabled."
		End If
	
		If Not(Window("Mobile Labs Trust Setup").WinButton("Next").GetROProperty("enabled")) Then
			strSummary = strSummary & " Next button is disabled."
			strActualResult = strActualResult & " Next button is disabled."
		Else
			blnResult = False
			strSummary = strSummary & " Next button is enabled."
			strActualResult = strActualResult & " Next button is enabled."
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
		strSummary = "EULA screen did not show up."
		arrStatus(intStep) = "Failed"
	
		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = "EULA screen is not being displayed."
		
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
		"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	
		Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
	End If
	
	'*********************************************************************************************************************
	
	intStep = intStep+1
	'*********************************************************************************************************************
	' Step3: Accept the License and hit 'Next' button.
	'Expected Result: Next button should get enabled. Destination Folder screen should come up, which has Change, Back, Next and Cancel buttons.
	
	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Accept the License and hit 'Next' button."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Accept the License and hit 'Next' button." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "Next button should get enabled. Destination Folder screen should come up, which has Change, " &_
	"Back, Next and Cancel buttons."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	'Accept the agreement
	Window("Mobile Labs Trust Setup").WinCheckBox("Accept EULA").Set "On"
	
	'Check if the Next button is enabled and if yes then hit it
	If  Window("Mobile Labs Trust Setup").WinButton("Next").GetROProperty("enabled") Then
		Window("Mobile Labs Trust Setup").WinButton("Next").Click
	
		'Check if the Destination Folder screen comes up
		If Window("Mobile Labs Trust Setup").Static("Destination Folder").GetROProperty("text") = "Destination Folder" Then
			strSummary = "Destination Folder screen is being displayed."
			strActualResult = "Destination Folder screen has been brought up."
			blnResult = True
		
			If Window("Mobile Labs Trust Setup").WinButton("Change").GetROProperty("enabled") Then
				strSummary = strSummary & " Change button is enabled."
				strActualResult = strActualResult & " Change button is enabled."
			Else
				blnResult = False
				strSummary = strSummary & " Change button is disabled."
				strActualResult = strActualResult & " Change button is disabled."
			End If
		
			If Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled") Then
				strSummary = strSummary & " Back button is enabled."
				strActualResult = strActualResult & " Back button is enabled."
			Else
				blnResult = False
				strSummary = strSummary & " Back button is disabled."
				strActualResult = strActualResult & " Back button is disabled."
			End If
		
			If Not(Window("Mobile Labs Trust Setup").WinButton("Next").GetROProperty("enabled")) Then
				strSummary = strSummary & " Next button is enabled."
				strActualResult = strActualResult & " Next button is enabled."
			Else
				blnResult = False
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
			strSummary = "Destination Folder screen didn't show up."
			arrStatus(intStep) = "Failed"
		
			ReDim Preserve arrActualResult(intStep)
			arrActualResult(intStep) = "Destination Folder screen is not being displayed."
			
			ReDim Preserve arrDetail(intStep)
			arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
			"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
		
			Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
		End If
		
	Else
		strSummary = "Next button is still disabled even after accepting the EULA."
		arrStatus(intStep) = "Failed"
	
		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = "EULA has been accepted but the Next button is disabled."
		
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
		"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	
		Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
	End If
	
	'*********************************************************************************************************************
	
	intStep = intStep+1
	'*********************************************************************************************************************
	' Step4: Keep the default installation location and hit "Next" button.
	'Expected Result: Ready to Install screen should be displayed with Back, Install and Cancel buttons enabled.
	
	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Keep the default installation location and hit 'Next' button."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Keep the default installation location and hit 'Next' button." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "Ready to Install screen should be displayed with Back, Install and Cancel buttons enabled."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	'Hit Next button to go to the next screen
	Window("Mobile Labs Trust Setup").WinButton("Next").Click
	
	'Check if the Ready to install screen is displayed
	If Window("Mobile Labs Trust Setup").Static("Ready").GetROProperty("text") = "Ready to install Mobile Labs Trust" Then
		strSummary = "Ready to install screen is being displayed."
		strActualResult = "Ready to install screen has been brought up."
		blnResult = True
	
		If Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled") Then
			strSummary = strSummary & " Back button is enabled."
			strActualResult = strActualResult & " Back button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Back button is disabled."
			strActualResult = strActualResult & " Back button is disabled."
		End If
	
		If Window("Mobile Labs Trust Setup").WinButton("Install").GetROProperty("enabled") Then
			strSummary = strSummary & " Install button is enabled."
			strActualResult = strActualResult & " Install button is enabled."
		Else
			blnResult = False
			strSummary = strSummary & " Install button is disabled."
			strActualResult = strActualResult & " Install button is disabled."
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
		strSummary = "Ready to install screen did not show up."
		arrStatus(intStep) = "Failed"
	
		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = "Ready to install screen is not being displayed."
		
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
		"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	
		Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
	End If
	
	'*********************************************************************************************************************
	
	intStep = intStep+1
	'*********************************************************************************************************************
	' Step5: Hit "Install" button.
	'Expected Result: Installing Mobile Labs Trust message should be displayed until installation completes. Once complete the Finish screen should appear 
	' with Finish button enabled and Back and Cancel buttons disabled.
	
	ReDim Preserve arrStepName(intStep)
	arrStepName(intStep) = "Hit 'Install' button."
	strStepsToReproduce = strStepsToReproduce & intStep & ": " & "Hit 'Install' button." & VBNewLine
	
	ReDim Preserve arrExpectedResult(intStep)
	arrExpectedResult(intStep) = "Installing Mobile Labs Trust message should be displayed until installation completes. Once " &_
	"complete the Finish screen should appear with Finish button enabled and Back and Cancel buttons disabled."
	
	ReDim Preserve arrStatus(intStep)
	arrStatus(intStep) = "Passed"
	
	'Hit Next button to go to the next screen
	Window("Mobile Labs Trust Setup").WinButton("Install").Click
	
	'Check if the EULA is brought up
	If Window("Mobile Labs Trust Setup").Static("Installing").GetROProperty("text") = "Installing Mobile Labs Trust" Then
	
		Do While Window("Mobile Labs Trust Setup").Static("Installing").Exist(1)
			' Wait until installation completes
		Loop
	
		If Window("Mobile Labs Trust Setup").Static("Completed").GetROProperty("text") = "Completed the Mobile Labs Trust Setup Wizard" Then
			strSummary = "Finish screen is being displayed."
			strActualResult = "Installation completion screen has been brought up."
			blnResult = True
		
			If Not(Window("Mobile Labs Trust Setup").WinButton("Back").GetROProperty("enabled")) Then
				strSummary = strSummary & " Back button is disabled."
				strActualResult = strActualResult & " Back button is disabled."
			Else
				blnResult = False
				strSummary = strSummary & " Back button is enabled."
				strActualResult = strActualResult & " Back button is enabled."
			End If
		
			If Window("Mobile Labs Trust Setup").WinButton("Finish").GetROProperty("enabled") Then
				strSummary = strSummary & " Install button is enabled."
				strActualResult = strActualResult & " Install button is enabled."
			Else
				blnResult = False
				strSummary = strSummary & " Install button is disabled."
				strActualResult = strActualResult & " Install button is disabled."
			End If
		
			If Not(Window("Mobile Labs Trust Setup").WinButton("Cancel").GetROProperty("enabled")) Then
				strSummary = strSummary & " Cancel button is disabled."
				strActualResult = strActualResult & " Cancel button is disabled."
			Else
				blnResult = False
				strSummary = strSummary & " Cancel button is enabled."
				strActualResult = strActualResult & " Cancel button is enabled."
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
			strSummary = "Finish screen did not show up."
			arrStatus(intStep) = "Failed"
		
			ReDim Preserve arrActualResult(intStep)
			arrActualResult(intStep) = "Installation completion screen is not being displayed."
			
			ReDim Preserve arrDetail(intStep)
			arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
			"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
		
			Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)
		End If
		
	Else
		strSummary = "Installing Mobile Labs Trust screen did not show up."
		arrStatus(intStep) = "Failed"
	
		ReDim Preserve arrActualResult(intStep)
		arrActualResult(intStep) = "Installing Mobile Labs Trust screen is not being displayed."
		
		ReDim Preserve arrDetail(intStep)
		arrDetail(intStep) = "Summary: " & strSummary & VBTab &_
		"Steps to Reproduce:" & VBNewLine & strStepsToReproduce
	
		Reporter.ReportEvent micFail, arrStepName(intStep), arrActualResult(intStep)	
	End If

'*********************************************************************************************************************

End If

'Write data into the TestResults excel file.
CreateTestResult arrStepName, arrExpectedResult, arrActualResult, arrDetail, arrStatus

''Close the installation wizard
Window("Mobile Labs Trust Setup").Close

Set objWScript = Nothing
Set objFSO = Nothing