'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiSwitch
' Test Description: Execute all MobiSwitch methods on  "switch" button.
' Steps:
' Step1 : Navigate to Switch object screen
' Step2 : Execute CaptureBitmap 
' Step3 : Execute CheckProperty
' Step4: Execute ChildObjects
' Step5: Execute Click with Bounadry Co-ordinates
'Setp6: Execute Click  with Random Co-ordinates
'Step7: Execute Click with Zero Co-ordinates
' Step8: Execute  Click without coordinates
' Step9: Execute Exist
' Step10: Execute GetROProperty
' Step11: Execute GetTOProperties
' Step12: Execute GetTOProperty
' Step13:Execute GetVisibleText
'Step14:Execute GetVisibleText without coordinates
' Step15 Execute RefreshObject
' Step16 Execute Set activate
' Step17 Execute Set activate
' Step18 Execute Set Deactivate
' Step19 Execute Set  Deactivate
' Step 20: Execute SetToProperty
' Step21: Execute ToString
' Step22: Execute WaitProperty
'##########################################################################################################
'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
'#######################################################

'#######################################################
'Initializations
'Initializations
intStep = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
'#######################################################
'*****************************************************************************************************************
' Step1: Navigate to Switch object screen
'Expected Result: Switch object screen should be displayed
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Navigate to Search screen" & VBNewLine
Environment("ExpectedResult") = "Switch object screen should be displayed"


'Set object for Switch

Set objMobiSwitch = MobiDevice("Phone Lookup").MobiSwitch("Switch")
strResult = Cstr( NavigateScreenOnPhoneLookup("Controls"  , objMobiSwitch , "Switch"))


' Step2:  Execute CaptureBitmap 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file to the defined location."
strResult = strResult & CStr(VerifyCaptureBitmap(objMobiSwitch))

' Step 3:  Execute 'CheckProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strValue = objMobiSwitch.GetROProperty("state")
strResult = strResult & CStr(VerifyCheckProperty(objMobiSwitch, "state", strValue , 5000, True))

'Step 4 : Execute ChildObjects
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
strResult = strResult & CStr(VerifyChildObjects(objMobiSwitch))

'Step 5 : Execute Click  with Boundary Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click with Boundary Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = strResult & CStr(VerifyClick(objMobiSwitch, "withboundarycoords"))

'Step 6 : Execute Click  with Random Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click with Random Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = strResult & CStr(VerifyClick(objMobiSwitch, "withrandomcoords"))

'Step 7 : Execute Click  with Zero Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click with Zero Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = strResult & CStr(VerifyClick(objMobiSwitch, "withzerovalues"))

'Step 8 : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = strResult & CStr(VerifyClick(objMobiSwitch, "withoutcoords"))


'Step 9 : Execute Exist
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = strResult & CStr(VerifyExist(objMobiSwitch, True, 5))

'Step 10: Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
strResult = strResult & CStr(VerifyGetROProperty(objMobiSwitch))

'Step 11: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("enabled")
strResult = strResult & CStr(VerifyGetTOProperties(objMobiSwitch, arrProps, strMissingProperties))

'Step 12: Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("enabled")
strResult = strResult & CStr(VerifyGetTOProperty(objMobiSwitch, arrProps, strNotFound))

'Step 13: Execute GetVisibleText
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
strResult = strResult & CStr(VerifyGetVisibleText(objMobiSwitch,True))


'Step 14: Execute GetVisibleText without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
strResult = strResult & CStr(VerifyGetVisibleText(objMobiSwitch,False))


'Step 15 : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = strResult & CStr(VerifyRefreshObject(objMobiSwitch))

'Step 16: Execute Set Activate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

strResult = strResult & CStr(VerifySet(objMobiSwitch, 1))

'Step 17: Execute Set Activate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

strResult = strResult & CStr(VerifySet(objMobiSwitch, 1))


'Step 18 : Execute Set deActivate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

strResult = strResult & CStr(VerifySet(objMobiSwitch, 0))


'Step 19 : Execute Set deActivate 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

strResult = strResult & CStr(VerifySet(objMobiSwitch, 0))



'Step 20: Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
strResult = strResult & CStr(VerifySetTOProperty(objMobiSwitch, arrProps, strSetFailedFor))

'Step 21: Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
strResult = strResult & CStr(VerifyTOString(objMobiSwitch))


'Step 22: Execute 'WaitProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strValue = objMobiSwitch.GetROProperty("state")

strResult = strResult & CStr(VerifyWaitProperty(objMobiSwitch, "state", strValue, 5000, True))
'*********************************************************************************************************************





