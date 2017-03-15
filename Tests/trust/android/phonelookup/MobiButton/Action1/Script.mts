

'##########################################################################################################
' Objective: Login to the PhoneLookup app and in the process test MobiButton  Methods
' Test Description: Execute all methods for MobiButton on SignIn button control 
'The methods are: CaptureBitmap, CheckProperty, ChildObjects, 
' Click, Exist, GetROProperty, GetTOProperties, GetTOProperty, GetVisibleText, RefreshObject, SetTOProperty, TOString, 
' WaitProperty

'Steps:
'Step 1:  Execute CaptureBitmap with .png extension
'Step 2:  Execute CaptureBitmap with .bmp extension
'Step 3:  Execute CaptureBitmap with .bmp extension already exist
'Step 4:  Execute CaptureBitmap with .png extension already exist
'Step 5:  Execute CheckProperty when object is visible
'Step 6:  Execute CheckProperty when object is not visible
'Step 7:  Execute ChildObjects 
'Step 8:  Execute Exist when object is visible
'Step 9:  Execute Exist when object is not visible
'Step 10:  Execute GetROProperty 
'Step 11:  Execute GetTOProperties 
'Step 12:  Execute GetTOProperty 
'Step 13:  Execute GetVisibleText  without coordinates
'Step 14:  Execute GetVisibleText  with coordinates
'Step 15:  Execute RefreshObject 
'Step 16:  Execute SetTOProperty 
'Step 17:  Execute ToString 
'Step 18:  Execute WaitProperty when object is visible
'Step 19:  Execute WaitProperty when object is not visible
'Step 20:  Execute Click  without coordinates
'Step 21:  Execute Click with  random coordinates
'Step 22:  Execute Click with boundary coordinates
'Step 23:  Execute Click with zero coordinates
'Step 24:  Execute Click with x co-ordinates
'Step 25:  Execute Click with y co-ordinates
'Step 26:  Execute Click with valid co-ordinates
'Step 27:  Execute Click with negative co-ordinates

'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName


CreateReportTemplate
'#######################################################

'#######################################################
'Initializations
intStep = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
Environment("StepName") = ""
Environment.Value("StepNo") = 0 
'#######################################################

'#######################################################
'Initial Setup

'Logout if a session is already in progress
Logout

'#######################################################

'Set object for Button
Set objMobiButton = MobiDevice("Phone Lookup").MobiButton("SignIn")
'initialize array for property and values
arrProperties = Array("text","id","Visible","buttontype","accessibilitylabel","resourcename")
arrPropertiesValue = Array("Sign In","2131230798","True","0","loginButton","loginButton")
arrProp =  Array("text","Visible","buttontype")
arrPropValue =  Array("Sign In","True","0")
' Step 1:  Execute CaptureBitmap with .png extension
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute CaptureBitmap on MobiButton with .png extension."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiButton." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot with .png extension and save the file to the defined location."
blnStepRC = VerifyCaptureBitmap(objMobiButton,"png")

' Step 2:  Execute CaptureBitmap with .bmp extension
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute CaptureBitmap on MobiButtonwith .bmp extension."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiButton." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot with .bmp extension and save the file to the defined location."
blnStepRC = VerifyCaptureBitmap(objMobiButton,"bmp")

' Step 3:  Execute CaptureBitmap with .bmp extension already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute CaptureBitmap on MobiButtonwith .bmp extension already exist."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Error message should be thrown for overriding the path"
blnStepRC = VerifyCaptureBitmap(objMobiButton,"override_bmp")

' Step 4:  Execute CaptureBitmap with .png extension already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute CaptureBitmap on MobiButtonwith .png extension already exist."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Error message should be thrown for overriding the path"
blnStepRC = VerifyCaptureBitmap(objMobiButton,"override_png")

'Step 5:  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute CheckProperty on MobiButton when object is visible."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
blnStepRC = VerifyCheckProperty(objMobiButton, "text", "Sign In", 5000, True)

'Step 6:  Execute ChildObjects recursively
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute ChildObjects on MobiButton recursively."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiButton." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children recursively (if any)."
blnStepRC = VerifyChildObjects(objMobiButton,"recursive",0)

'Step 6:  Execute ChildObjects non-recursively
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute ChildObjects on MobiButton non-recursively."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiButton." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children non-recursively (if any)."
blnStepRC = VerifyChildObjects(objMobiButton,"nonrecursive",0)

'Step 7:  Execute Exist when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Exist on MobiButton when object is visible."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return true."
blnStepRC = VerifyExist(objMobiButton, True, 5)

''Step 8:  Execute GetROProperty 
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute GetROProperty on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnStepRC = VerifyGetROProperty(objMobiButton,arrProperties,arrPropertiesValue)

'Step 9:  Execute GetTOProperties 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute GetTOProperties on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiButton." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
blnStepRC = VerifyGetTOProperties(objMobiButton, arrProp)

'Step 10:  Execute GetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute GetTOProperty on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
blnStepRC = VerifyGetTOProperty(objMobiButton,arrProp,arrPropValue)
'
'Step 13:  Execute GetVisibleText  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
'Environment.Value("StepNo")  = Environment.Value("StepNo")  + 1
Environment.Value("Description") = "Execute GetVisibleText on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiButton." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnStepRC = VerifyGetVisibleText(objMobiButton,False)

'Step 14:  Execute GetVisibleText  with coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
'Environment.Value("StepNo")  = Environment.Value("StepNo")  + 1
Environment.Value("Description") = "Execute GetVisibleText on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiButton." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnStepRC = VerifyGetVisibleText(objMobiButton,True)

'Step 15:  Execute RefreshObject 
'#######################################################  
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute RefreshObject on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiButton." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnStepRC = VerifyRefreshObject(objMobiButton)

'Step 16:  Execute SetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute SetTOProperty on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnStepRC =VerifySetTOProperty(objMobiButton, arrProperties)

'Step 17:  Execute ToString 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute TOString on MobiButton."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiButton." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnStepRC = VerifyTOString(objMobiButton)

'Step 18:  Execute WaitProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute WaitProperty on MobiButton when object is visible."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnStepRC = VerifyWaitProperty(objMobiButton, "buttontype", 0, 5000, True)

Login  "mobilelabs" , "demo"
Set objMobiButton  = MobiDevice("Phone Lookup").MobiButton("Search")

'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton without coordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiButton, "withoutcoords")

GoToScreen "Search"

'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton for random co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiButton, "withrandomcoords")

GoToScreen "Search"

'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton for boundary co-ordinates at Top-Left corner."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left corner."
blnStepRC = VerifyClick(objMobiButton, "withboundarycoordsTopLeft")

GoToScreen "Search"

'Step 22:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton for boundary co-ordinates at Top-Right corner."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right corner."
blnStepRC = VerifyClick(objMobiButton, "withboundarycoordsTopRight")

GoToScreen "Search"

'Step 22:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton for boundary co-ordinates at Bottom-Left corner."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left corner."
blnStepRC = VerifyClick(objMobiButton, "withboundarycoordsBottomLeft")

GoToScreen "Search"

'Step 22:  Execute Click with boundary coordinates at Bottom-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton for boundary co-ordinates at Bottom-Right corner."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right corner."
blnStepRC = VerifyClick(objMobiButton, "withboundarycoordsBottomRight")

GoToScreen "Search"

'Step 23:  Execute Click with zero coordinates
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment.Value("Description") = "Execute Click on MobiButton for zero co-ordinates."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiButton." & VBNewLine
'Environment("ExpectedResult") = "Click should work correctly with zero co-ordinates."
'blnStepRC = VerifyClick(objMobiButton, "withzerovalues")

GoToScreen "Search"


'Step 24:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton with x co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiButton, "withxvalue")

GoToScreen "Search"

'Step 25:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton with y co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiButton, "withyvalue")

GoToScreen "Search"


'Step 26:  Execute Click with valid co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton with  valid co-ordinates."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with  valid co-ordinates."
blnStepRC = VerifyClick(objMobiButton, "withvalidvalue")

GoToScreen "Search"

'Step 27:  Execute Click with negative co-ordinates
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment.Value("Description") = "Execute Click on MobiButton with  negative co-ordinates."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiButton." & VBNewLine
'Environment("ExpectedResult") = "Click should throw run time error."
'blnStepRC = VerifyClick(objMobiButton, "withnegativecoords")
wait 4


LogOut

'Navigate to controls screen to verify object  property when not visible
NavigateScreenOnPhoneLookup  "Controls"  , MobiDevice("Phone Lookup").MobiDatetimePicker("DatePicker") ,  "DatePicker"

Set objMobiButton = MobiDevice("Phone Lookup").MobiButton("SignIn")

'Step 19:  Execute WaitProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute WaitProperty on MobiButton when object is not visible."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnStepRC = VerifyWaitProperty(objMobiButton, "buttontype", 0, 5000, False)

'Step 9:  Execute Exist when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Exist on MobiButton when object is not visible."
'@click the button so the button will not be visible
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return false."
blnStepRC = VerifyExist(objMobiButton, False, 5)

'Step 6:  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute CheckProperty on MobiButton when object is not visible."
'@click the button so the button will not be visible
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiButton." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
blnStepRC = VerifyCheckProperty(objMobiButton, "text", "Sign In", 5000, False)
''*********************************************************************************************************************
EndTestIteration









