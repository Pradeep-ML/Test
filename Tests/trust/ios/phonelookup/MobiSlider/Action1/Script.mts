'##########################################################################################################
' Objective: Login to the PhoneLookup app if the slider object is not visible and test MobiSlider object methods
' Test Description: Execute all methods for MobiSlidet  in the UISlider from the controls list
' Expected Result: All Mobislider methods should work correctly. The methods are: CaptureBitmap , Check , CheckProperty , ChildObjects , Click
'GetPercentage , GetROProperty , GetTOProperties , GetTOProperty , Output 'RefreshObject , Set , SetTOProperty , ToString , WaitProperty , Exist

'#######################################################
'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
'#######################################################
Set objMobiSlider = MobiDevice("PhoneLookup").MobiSlider("Slider")

'#######################################################
'Initializations
intStep = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
Environment("intStepNo") = 0
'#######################################################
'Create an html report template
CreateReportTemplate()
'#######################################################
'Initial Setup

'Set object for MobiSlider
Set objMobiSlider = MobiDevice("PhoneLookup").MobiSlider("Slider")

'#######################################################

' Step1: Navigate to UISlider in the control list in PhoneLookUp
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Navigate to UISlider page of PhoneLookUp App."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Navigate to UISlider page of PhoneLookUp App." & VBNewLine
Environment("ExpectedResult") = "User should be navigated to Slider Page"

		'Logout 
		LogOut
		'Login and navigate to UISlider page
		strResult = LoginAndNavigateToControlsPage("UISlider", objMobiSlider)

' Step2: Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSlider with .png format."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Image should get captured in .png format."
strResult = VerifyCaptureBitmap(objMobiSlider , "png")

' Step3: Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSlider with .bmp format."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Image should get captured in .bmp format."
strResult = VerifyCaptureBitmap(objMobiSlider , "bmp")

' Step4: Execute CaptureBitmap to overwrite an image with .png format
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSlider to override an image with  .png  format."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Proper error message should be thrown"
strResult = VerifyCaptureBitmap(objMobiSlider , "override_png")

' Step5: Execute CaptureBitmap to overwrite an image with .bmp format
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSlider to override an image with  .bmp  format."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Proper error message should be thrown"
strResult = VerifyCaptureBitmap(objMobiSlider , "override_bmp")

' Step6: Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckPropertywhen object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSlider when object is visile." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."

'strPropertyValue = objMobiSlider.GetROProperty("visible")
strResult = VerifyCheckProperty(objMobiSlider, "visible","True", 5000, True)


' Step8: Execute Click without co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSlider without co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
Set objMobiSlider = MobiDevice("PhoneLookup").MobiSlider("Slider")
objMobiSlider.Set 90
strResult = VerifyClick(objMobiSlider, "withoutcoords")


'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiWebButton without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withoutcoords")

'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withrandomcoords")
'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withboundarycoordsTopLeft")


'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withboundarycoordsTopRight")


'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withboundarycoordsBottomLeft")


'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withboundarycoordsBottomRight")


'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withxvalue")



'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
objMobiSlider.Set 90
blnStepRC = VerifyClick(objMobiSlider, "withyvalue")

' Step 16: Execute Exist when object is  visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Exist should return True"
strResult = VerifyExist(objMobiSlider, True, 5)


' Step 17: Execute GetROProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
arrProp = Array("accessibilityidentifier" , "accessibilitylabel" , "percentage")
appPropVal = Array("" , "" , "100")
objMobiSlider.Set 100
objMobiSlider.WaitProperty "percentage", 100 , 5000
objMobiSlider.WaitProperty "value", 100 , 5000

strResult = VerifyGetROProperty(objMobiSlider , arrProp , appPropVal) 

' Step 18: Execute GetTOProperties
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("id", "minvalue", "maxvalue")
strResult = VerifyGetTOProperties(objMobiSlider, arrProps)

' Step 19: Execute GetTOProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("id",  "minvalue", "maxvalue")
arrPropValues = Array("1","0","100")
strResult = VerifyGetTOProperty(objMobiSlider, arrProps, arrPropValues)

' Step 20: Execute Set with percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  with percentage"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Set should set the passed percentage correctly."
Set objMobiSlider = MobiDevice("PhoneLookup").MobiSlider("Slider")
strResult = VerifySet(objMobiSlider, 60 , "")

''Step 24 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSlider , "withcentervalues" , "notoccluded")
''#############################################################
'
''Step 23 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSlider , "withoutcoords" , "notoccluded")
''#############################################################

'' Step 21: Execute Set without percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Set  without percentage."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSlider." & VBNewLine
'Environment("ExpectedResult") = "Set should throw correct error message."
'strResult = VerifySet(objMobiSlider, null , "")

' Step 22: Execute Set with negative percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  with negative percentage"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Set should throw correct error message."
strResult = VerifySet(objMobiSlider, -70 , "")

' Step 23: Execute SetTOProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetTOProperty on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
strResult = VerifySetTOProperty(objMobiSlider, arrProps)


' Step 24: Execute ToString
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute TOString on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "ToString should return the object type and class."
strResult = VerifyToString(objMobiSlider)

' Step 25: Execute WaitProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiSlider , "visible",True, 5000, True)


' Step 27: Execute GetPercentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetPercentage on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetPercentage  on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetPercentage should return the percentage the slider is set on"
'strDeviceType = MobiDevice("PhoneLookup").GetROProperty("devicetype")
strResult = VerifyGetPercentage(objMobiSlider)

'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiSlider, "nonrecursive" , 3 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiSlider, "recursive" ,3)


' Step 29: Execute RefreshObject
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiSlider)

'Logout  from Slider screen
LogOut 

' Step7: Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSlider when object is not visile." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strResult = VerifyCheckProperty(objMobiSlider, "visible","True", 2000, False)

' Step 16: Execute Exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Exist should return False."
strResult = VerifyExist(objMobiSlider, False , 5)

' Step 26: Execute WaitProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiSlider , "visible", True, 2000, False)
'
''**************************************************************************************************************

EndTestIteration







