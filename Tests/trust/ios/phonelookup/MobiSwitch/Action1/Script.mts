'##########################################################################################################
' Objective: Login to the PhoneLookup app and in the process test MobiSwitch  Methods
' Test Description: Execute all methods for MobiSwitch on Remember Me Switch control
'The methods are: CaptureBitmap, CheckProperty, ChildObjects, 
' Click, Exist, GetROProperty, GetTOProperties, GetTOProperty, GetVisibleText, RefreshObject, SetTOProperty, TOString, 
' WaitProperty , Set

'######################################################
'Declare Variables

Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
CreateReportTemplate
'#######################################################


'Initializations
intStep = 0
Environment("intStepNo") = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
Environment("StepName") = ""
'#######################################################

'Initial Setup

'Logout if a session is already in progress
Logout

'#######################################################


'Set object for Switch
Set objMobiSwitch = MobiDevice("PhoneLookup").MobiSwitch("RememberMe")
arrTOProperties = Array("enabled")
arrTOPropertiesValue = Array("True")
arrROProperties = Array("enabled","name","nativeclass")
arrROPropertiesValue = Array("True","Switch","UISwitch")


' Step 1:  Execute CaptureBitmap with .png extension
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot with .png extension and save the file to the defined location."
blnStepRC = VerifyCaptureBitmap(objMobiSwitch,"png")

' Step 2:  Execute CaptureBitmap with .bmp extension
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .bmp extension."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot with .bmp extension and save the file to the defined location."
blnStepRC = VerifyCaptureBitmap(objMobiSwitch,"bmp")

' Step 3:  Execute CaptureBitmap with .bmp extension already exist
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .bmp extension already exist."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Error message should occur for override the file"
blnStepRC = VerifyCaptureBitmap(objMobiSwitch,"override_bmp")

' Step 4:  Execute CaptureBitmap with .png extension already exist
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .png extension already exist."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Error message should occur for override the file"
blnStepRC = VerifyCaptureBitmap(objMobiSwitch,"override_png")

' Step 5:  Execute CheckProperty when object exist
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
'strValue = objMobiSwitch.GetROProperty("state")
'If LCase(strValue) = "on" Then
'	objMobiSwitch.Set eDEACTIVATE
'	strValue = "Off"
'ElseIf  LCase(strValue) = "off" Then
'	objMobiSwitch.Set eACTIVATE
'	strValue = "On"
'End If
blnStepRC = VerifyCheckProperty(objMobiSwitch, "visible", "True" , 5000, True)

'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiSwitch, "nonrecursive" , 3 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiSwitch, "recursive" ,3)


'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSwitch without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
Set objMobiSwitch = MobiDevice("PhoneLookup").MobiSwitch("RememberMe")
blnStepRC = VerifyClick(objMobiSwitch, "withoutcoords")


'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiSwitch, "withrandomcoords")


'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiSwitch, "withboundarycoordsTopLeft")

'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiSwitch, "withboundarycoordsTopRight")


'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiSwitch, "withboundarycoordsBottomLeft")

'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiSwitch, "withboundarycoordsBottomRight")


'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiSwitch, "withxvalue")


'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSwitch with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiSwitch, "withyvalue")


' Step 16:  Execute Exist  when object is visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiSwitch when object is visible."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return true."
blnStepRC = VerifyExist(objMobiSwitch, True, 5)

''Step 24 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSwitch , "withcentervalues" , "notoccluded")
''#############################################################
'
''Step 23 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSwitch , "withoutcoords" , "notoccluded")
''#############################################################
'
'' Step 18:  Execute GetROProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnStepRC = VerifyGetROProperty(objMobiSwitch, arrROProperties, arrROPropertiesValue)

' Step19:  Execute GetTOProperties 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
blnStepRC = VerifyGetTOProperties(objMobiSwitch,arrTOProperties)

' Step 20:  Execute GetTOProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
blnStepRC =  VerifyGetTOProperty(objMobiSwitch, arrTOProperties, arrTOPropertiesValue)

' Step 21:  Execute GetVisibleText  on deactivated switch without coordinates 
'################################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
objMobiSwitch.Set eDEACTIVATE
blnStepRC = VerifyGetVisibleText(objMobiSwitch, False)

 'Step 22:  Execute GetVisibleText  on activated switch without coordinates 
'################################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiSwitch on activated switch without coordinates ."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
objMobiSwitch.Set  eACTIVATE
blnStepRC = VerifyGetVisibleText(objMobiSwitch, False)

 'Step 23:  Execute GetVisibleText  on deactivated switch with coordinates 
'################################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiSwitch on deactivated switch with coordinates ."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
objMobiSwitch.Set   eDEACTIVATE
blnStepRC = VerifyGetVisibleText(objMobiSwitch, True)

 'Step 24:  Execute GetVisibleText  on activated switch with coordinates 
'################################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiSwitch on activated switch with coordinates ."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
objMobiSwitch.Set   eACTIVATE
blnStepRC = VerifyGetVisibleText(objMobiSwitch, True)

' Step 25:  Execute RefreshObject 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject  on MobiSwitch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnStepRC = VerifyRefreshObject(objMobiSwitch)

' Step 26:  Execute Set  to  activate 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  on MobiSwitch to  activate ."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."
blnStepRC = VerifySet( objMobiSwitch, 1,null)

' Step 27:  Execute Set  to  activate  an activated switch
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  on MobiSwitch to  activate  an activated switch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."
blnStepRC = VerifySet(objMobiSwitch, 1,null)


' Step 28:  Execute Set  to  deactivate an activated switch
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  on MobiSwitch to  deactivate  an activated switch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: DeActivated."
blnStepRC = VerifySet(objMobiSwitch, 0,null)

' Step 29:  Execute Set  to  deactivate  a deactivated switch
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  on MobiSwitch to  deactivate  an deactivated switch."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: DeActivated."
blnStepRC = VerifySet(objMobiSwitch, 0,null)


'' Step 30:  Execute Set  without parameter
''#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Set  on MobiSwitch without parameter."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
'Environment("ExpectedResult") = "Error message should displayed for no parameter."
'blnStepRC = VerifySet(objMobiSwitch, null ,null)


' Step 31:  Execute SetTOProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetTOProperty on MobiSwitch"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnStepRC = VerifySetTOProperty(objMobiSwitch,arrTOProperties)

' Step 32:  Execute TOString 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute TOString on MobiSwitch"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnStepRC = VerifyTOString(objMobiSwitch)

' Step 33:  Execute WaitProperty when object is visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiSwitch when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnStepRC = VerifyWaitProperty(objMobiSwitch, "visible", "True", 5000, True)

'Logout


GoToScreeniOS  "Controls" 
' Step 34:  Execute WaitProperty when object is not visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiSwitch when object is not visible"

Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnStepRC = VerifyWaitProperty(objMobiSwitch, "visible", True, 5000, False)

' Step 6:  Execute CheckProperty when object doesn't exist
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiSwitch when object doesn't exist."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
blnStepRC = VerifyCheckProperty(objMobiSwitch, "visible", "True" , 5000, False)

' Step 17:  Execute Exist  when object is not visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiSwitch when object is not visible."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return false."
blnStepRC = VerifyExist(objMobiSwitch, False, 5)
'**********************************************************************************************************************************************************************


EndTestIteration





