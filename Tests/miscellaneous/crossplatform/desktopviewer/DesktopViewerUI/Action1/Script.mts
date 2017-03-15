'##########################################################################################################
'Objective: Launch the Desktop Viewer via nativeautomation and test the UI
' Test Description: Check all dialogs and menu options on the Desktop Viewer
'##########################################################################################################

'Option Explicit

'Declare Variables
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
Dim objWindow

'#######################################################

'#######################################################
'Initializations
intStep = 0
Environment("Component") = "Desktop Viewer"
'#######################################################
'Input parameters
Set objWindow  =  WpfWindow("DesktopViewer")
'Set objMenu = WpfWindow("DesktopViewer").WpfMenu("MenuMainTop")

'Create an html report template
CreateReportTemplate()
'#######################################################

' Step1: Check all menu items except Debug menu
'Expected Result: All menu options should be displayed properly
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check all menu items except Debug menu"
Environment("ExpectedResult") = "All menu options should be displayed properly"

blnResult = VerifyDesktopViewerMenuOptions(objWindow,False)

'''*********************************************************************************************************************
' Step2:  Check all menu items along with Debug menu
'Expected Result: All menu options should be displayed properly
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check all menu items along with Debug menu"
Environment("ExpectedResult") = "All menu options should be displayed properly"
blnResult = VerifyDesktopViewerMenuOptions(objWindow,True)

' Step3:  Execute CaptureBitmap with .bmp format
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify method for capturing .bmp file"
'Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
'blnResult = VerifyCaptureBitmap(objMobiWebButton , "bmp")

' Step 4:  Execute CaptureBitmap to override .bmp file
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = " Verify override message for already existing .bmp  file"
'Environment("ExpectedResult") = "CaptureBitMap should display override message"
'blnResult = VerifyCaptureBitmap(objMobiWebButton , "override_bmp")


'End test iteration
EndTestIteration()
