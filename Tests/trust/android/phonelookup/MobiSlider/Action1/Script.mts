'##########################################################################################################
' Objective: Login to the PhoneLookup app if the slider object is not visible and test MobiSlider object methods
' Test Description: Execute all methods for MobiSlidet 
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
Set objMobiSlider = MobiDevice("Phone Lookup").MobiSlider("Slider")

'#######################################################

' Step1: Call function to navigate to Seekbar screen.
'Expected Result : User should be navigated to Seekbar screen.
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Call function to navigate to Seekbar screen.." & VBNewLine
Environment("ExpectedResult") = "User should be navigated to Seekbar Screen"

Set objMobiSlider = MobiDevice("Phone Lookup").MobiSlider("Slider")

		'Call function to navigate to SeekBar screen
		NavigateScreenOnPhoneLookup "Controls" , objMobiSlider , "SeekBar"   
	
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

' Step4: Execute CaptureBitmap to override an image with .png format
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSlider to override an image with  .png  format."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Proper error message should be thrown."
strResult = VerifyCaptureBitmap(objMobiSlider , "override_png")

' Step5: Execute CaptureBitmap to override an image with .bmp format
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSlider to override an image with  .bmp  format."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Proper error message should be thrown."
strResult = VerifyCaptureBitmap(objMobiSlider , "override_bmp")

' Step6: Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiSlider when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSlider when object is visile." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and return True"

'strPropertyValue = objMobiSlider.GetROProperty("visible")
strResult = VerifyCheckProperty(objMobiSlider, "visible", True , 6000, True)

' Step8: Execute Click without co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSlider without co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
Set objMobiSlider = MobiDevice("Phone Lookup").MobiSlider("Slider")
objMobiSlider.Set 90
strResult = VerifyClick(objMobiSlider, "withoutcoords")

' Step9: Execute Click with random co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSlider with random co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 90
strResult = VerifyClick(objMobiSlider, "withrandomcoords")

' Step 13:  Execute Click  with boundary coordinates at Top-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 50
Wait 2
blnFlag = VerifyClick(objMobiSlider, "withboundarycoordsTopLeft")

' Step 13:  Execute Click  with boundary coordinates at Top-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Right corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 50 
Wait 2
blnFlag = VerifyClick(objMobiSlider,"withboundarycoordsTopRight")

' Step 13:  Execute Click  with boundary coordinates at Bottom-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Left corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 50 
Wait 2
blnFlag = VerifyClick(objMobiSlider,"withboundarycoordsBottomLeft")

' Step 13:  Execute Click  with boundary coordinates at Bottom-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Right corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 50
Wait 2
blnFlag = VerifyClick(objMobiSlider,"withboundarycoordsBottomRight")


' Step12: Execute Click with x co-ordinate
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSlider with x co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 90
strResult = VerifyClick(objMobiSlider, "withxvalue")

' Step 13: Execute Click with y co-ordinate
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSlider with y co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 90
strResult = VerifyClick(objMobiSlider, "withyvalue")

' Step 14: Execute Click with valid co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSlider with valid co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiSlider.Set 0
wait 2
strResult = VerifyClick(objMobiSlider, "withvalidvalue")


' Step 16: Execute Exist when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiSlider  when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Exist should return True"
strResult = VerifyExist(objMobiSlider, True, 5000)


' Step 17: Execute GetROProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiSlider"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."

arrProp = Array("minvalue" , "maxvalue")
arrPropVal = Array(0,100)
strResult = VerifyGetROProperty(objMobiSlider , arrProp , arrPropVal)

' Step 18: Execute GetTOProperties
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("minvalue", "maxvalue")
strResult = VerifyGetTOProperties(objMobiSlider, arrProps)

' Step 19: Execute GetTOProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("minvalue", "maxvalue" , "id")
arrPropValues = Array(0,100 , 2131230878)
strResult = VerifyGetTOProperty(objMobiSlider, arrProps, arrPropValues)

' Step 20: Execute Set with percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set on MobiSlide with valid percentage."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Set should set the passed percentage correctly."
'Set objMobiSlider = MobiDevice("Phone Lookup").MobiSlider("Slider")
strResult = VerifySet(objMobiSlider, 60 , "")


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
Environment("Description") = "Execute WaitProperty on MobiSlider when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and should return True"
strResult = VerifyWaitProperty(objMobiSlider , "visible" , True , 5000 , True)

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

' Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiSlider recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiSlider.” & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiSlider ,"recursive",0)

' Step 7:  Execute  ChildObjects nonrecusrively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiSlider non-recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiSlider ,"nonrecursive",0)


' Step 29: Execute RefreshObject
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject on MobiSlider."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiSlider)


LogOut

' Step 16: Execute Exist when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiSlider  when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "Exist should return False"
strResult = VerifyExist(objMobiSlider, False , 5000)


' Step7: Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiSlider when object is not visible "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSlider when object is not visile." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and return False."
strResult = VerifyCheckProperty(objMobiSlider, "visible", True, 6000, False)


' Step 26: Execute WaitProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiSlider when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSlider." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and should  return False."
strResult = VerifyWaitProperty(objMobiSlider , "visible" , True , 5000 , False)

''**************************************************************************************************************

EndTestIteration





















