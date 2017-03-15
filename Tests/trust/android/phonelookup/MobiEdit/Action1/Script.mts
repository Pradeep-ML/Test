
'##########################################################################################################
'Objective: Login to the PhoneLookup app and test MobiEdit
' Test Description: Execute all MobiEdit methods on  Username Edit Box
'##########################################################################################################

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
Environment("MethodName")  = ""
Environment("intStepNo") = 0
Environment("Status") = ""
'#######################################################

'Input values
arrTOProps = Array("visible" , "defaultvalue" )
arrToPropValues = Array( True , "Enter Username" )

arrROProps = Array("nativeclass" ,  "enabled")
arrROPropValues = Array("android.widget.EditText" , "True")

'Create an html report template
CreateReportTemplate()

'#######################################################
' Step: Navigate to Search screen
'Expected Result: Search screen should be displayedintStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Navigate to Search screen" & VBNewLine
Environment("ExpectedResult") = "Search screen should be displayed"

'Set object for Button
Set objMobiEdit =MobiDevice("Phone Lookup").MobiEdit("UserName")


'Call navigate to screen function 
 NavigateScreenOnPhoneLookup  "Login"  , objMobiEdit   , ""

'Capture Sign In button screenshot  before execution

strImagePathBefore = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathBefore , True

'###########################################################

' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap with .png format" & VBNewLine
Environment("Description") = "Execute method to capture image in .png format"
Environment("ExpectedResult") = "Image should get captured in .png format"
blnResult = VerifyCaptureBitmap(objMobiEdit , "png")

' Step:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap with .bmp format" & VBNewLine
Environment("Description") = "Execute method to capture image in .bmp format"
Environment("ExpectedResult") = "Image should get captured in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiEdit , "bmp")

' Step :  Execute CaptureBitmap to override an .bmp image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap  to override an .bmp image" & VBNewLine
Environment("Description") = "Execute method  to override an .bmp image"
Environment("ExpectedResult") = "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiEdit , "override_bmp")

' Step :  Execute CaptureBitmap to override an .png image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap  to override an .png image" & VBNewLine
Environment("Description") = " Execute method  to override an .png image"
Environment("ExpectedResult") =  "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiEdit , "override_png")

' Step :  Execute  CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute  CheckProperty when object is visible" & VBNewLine
Environment("Description") = "Execute method to check property value when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiEdit, "id", 2131230792 , 5000, True)

'Navigate to other screen
Login "mobilelabs" , "demo"

' Step :  Execute  CheckProperty when object is not  visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute  CheckProperty when object is not visible" & VBNewLine
Environment("Description") = " Execute method to check property value when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiEdit, "id", 2131230792 , 5000, False)

'Navigate back to object screen
Logout

'Step  : Execute ChildObjects recursively
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("Description") ="Verfiy child object count recursively"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) else 0"
blnResult = VerifyChildObjects(objMobiEdit,"recursive",0)


'Step  : Execute ChildObjects non-recursively
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("Description") ="Verfiy child object count non-recursively"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) else 0"
blnResult = VerifyChildObjects(objMobiEdit,"nonrecursive",0)

'Step  : Execute Clear  on long string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Clear on MobiEdit." & VBNewLine
Environment("Description") ="Verfiy method to clear long string"
Environment("ExpectedResult") = "There should be no text in the Edit box"
blnResult = VerifyClear( objMobiEdit , "withlongtext")

'Step  : Execute Clear when no text  is visible in Edit box
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Clear on MobiEdit." & VBNewLine
Environment("Description") =" Execute Clear when no text  is visible in Edit box"
Environment("ExpectedResult") = "There should be no text in the Edit box"
blnResult = VerifyClear( objMobiEdit , "withnotext")

'Step  : Execute Clear  on small  string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Clear on MobiEdit." & VBNewLine
Environment("Description") ="Verfiy method to clear small  string"
Environment("ExpectedResult") = "There should be no text in the Edit box"
blnResult = VerifyClear( objMobiEdit , "withtext")

'Step  Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiEdit when object is visible" & VBNewLine
Environment("Description") = "Verify method when object is visible"
Environment("ExpectedResult") = "Exist should return True when object is visible"
blnResult = VerifyExist(objMobiEdit, True, 5)

'Navigate to other screen
Login "mobilelabs" , "demo"

'Step  Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiEdit when object is visible" & VBNewLine
Environment("Description") = " Verify method when object is not visible"
Environment("ExpectedResult") = "Exist should return False when object is not visible."
blnResult = VerifyExist(objMobiEdit, False, 10)

'Navigate back to object screen
LogOut

'Step  Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiEdit." & VBNewLine
Environment("Description") = "Verify object run time values"
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnResult = VerifyGetROProperty(objMobiEdit , arrROProps , arrROPropValues)

'Step  : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiEdit." & VBNewLine
Environment("Description") = "Verify object description properties"
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
blnResult = VerifyGetTOProperties(objMobiEdit, arrTOProps)

'Step : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiEdit." & VBNewLine
Environment("Description") = " Verify object description propertie and their values"
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
blnResult =  VerifyGetTOProperty(objMobiEdit, arrTOProps, arrToPropValues)

'Step  : Execute GetVisibleText method with coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute GetVisibleText on MobiEdit with coordinates." & VBNewLine
Environment("Description") = " Execute method with co-ordinates"
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnResult = VerifyGetVisibleText(objMobiEdit, True)

'Step : Execute GetVisibleText method without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute GetVisibleText on MobiEdit without coordinates." & VBNewLine
Environment("Description") = " Execute method without  co-ordinates"
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnResult = VerifyGetVisibleText(objMobiEdit, False)


'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiEdit." & VBNewLine
Environment("Description") = "Verify object refresh"
Environment("ExpectedResult") = "RefreshObject should re-identify  the object in the application"
blnResult = VerifyRefreshObject(objMobiEdit)

'Step  : Execute Set  with Alphanumeric String
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set" & VBNewLine
Environment("Description") = "Execute method to set alphanumeric string"
Environment("ExpectedResult") = "User should be able to set value"
blnResult = VerifySet(objMobiEdit , "testing123" , null)

'Step  : Execute Set  with  String cointaining special characters
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set" & VBNewLine
Environment("Description") = "Execute method to set  string cointaining special characters"
Environment("ExpectedResult") = "User should be able to set value"
blnResult = VerifySet(objMobiEdit ,  "Testing..~!@#$%^&*()_+{}|:<>?/.,';\][=-`0123456789" , null)

'Step  : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiEdit." & VBNewLine
Environment("Description") = "Verify object  type and class"
Environment("ExpectedResult") = "ToString should return the object type and class."
blnResult = VerifyTOString(objMobiEdit)

'Step   : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiEdit when object is visible." & VBNewLine
Environment("Description") = " Verify waitproperty when object is visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value"
blnResult = VerifyWaitProperty(objMobiEdit, "id", 2131230792, 5000, True)

'Navigate to other screen
Login "mobilelabs" , "demo"

'Step  : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiEdit when object is not  visible." & VBNewLine
Environment("Description") = "Verify waitproperty when object is not visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and return False"
blnResult = VerifyWaitProperty(objMobiEdit, "id", 2131230792 , 15000, False)

'Navigate back to object screen
LogOut

'Step  : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiEdit." & VBNewLine
Environment("Description") = "Property values should get updated"
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnResult = VerifySetTOProperty(objMobiEdit, arrTOProps)

'*********************************************************************************************************************

objMobiEdit.Set  "mobilelabs"

'Step  : Execute Click with  Boundary coordinates at Top-Left corner
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with Boundary coordinates." & VBNewLine
Environment("Description") = "Execute method with boundary co-ordinates at Top-Left corner"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiEdit, "withboundarycoordsTopLeft")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True
If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click with  Boundary coordinates at Top-Right corner
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with Boundary coordinates." & VBNewLine
Environment("Description") = "Execute method with boundary co-ordinates at Top-Right corner"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiEdit, "withboundarycoordsTopRight")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True
If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click with  Boundary coordinates at Bottom-Left corner
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with Boundary coordinates." & VBNewLine
Environment("Description") = "Execute method with boundary co-ordinates at Bottom-Left corner"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiEdit, "withboundarycoordsBottomLeft")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True
If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click with  Boundary coordinates at Bottom-Right corner
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with Boundary coordinates." & VBNewLine
Environment("Description") = "Execute method with boundary co-ordinates at Bottom-Right corner"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiEdit, "withboundarycoordsBottomRight")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True
If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with Random coordinates." & VBNewLine
Environment("Description") = " Execute method with Random co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiEdit, "withrandomcoords")


'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute Click with  Zero coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with Zero coordinates." & VBNewLine
Environment("Description") = "Execute method with zero co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiEdit, "withzerovalues")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit without coordinates." & VBNewLine
Environment("Description") = "Execute method without  co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiEdit, "withoutcoords")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click  at negative co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with negative coordinates." & VBNewLine
Environment("Description") = " Execute method with negative co-ordinates"
Environment("ExpectedResult") = "Click should throw error message"
blnResult =  VerifyClick(objMobiEdit, "withnegativecoords")


'Step  : Execute Click  at only one co-ordinate (Only X)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with only X coordinate" & VBNewLine
Environment("Description") = " Execute method with only X co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiEdit, "withxvalue")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click  at only one co-ordinate (Only Y)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiEdit with only Y coordinate" & VBNewLine
Environment("Description") = "Execute method with only Y co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiEdit, "withyvalue")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step  : Execute Click  at  any valid value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click  at  any valid value." & VBNewLine
Environment("Description") = "Execute method with any valid co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiEdit, "withvalidvalue")

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute DblClick  without  co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = "Execute method without co-ordinates"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withoutcoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute DblClick  with boundary co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = " Execute method with Boundary co-ordinates"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withboundarycoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute DblClick  with Random co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = "Execute method with Random co-ordinates"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withrandomcoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute DblClick  with Zero co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = " Execute method with Zero co-ordinates"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withzercoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute DblClick  with Negative  co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = "Execute method with Negative  co-ordinates"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withnegativecoords") 

'Step : Execute DblClick  with only X  co-ordinate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = "Execute method with only X Co-ordinate"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withonlyxcoord") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute DblClick  with only Y  co-ordinate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick ." & VBNewLine
Environment("Description") = " Execute method with only Y  Co-ordinate"
Environment("ExpectedResult") = "DblClick  should work correctly."
blnResult =  VerifyDblClick(objMobiEdit , "withonlyycoord") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  without  co-ordinates with valid lapse time
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = "Execute method without co-ordinates with valid lapse time"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "validlapsewithoutcoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with boundary co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with Boundary co-ordinates with valid lapse time"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "validlapseboundarycoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with Random co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with Random co-ordinates and valid lapse time"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "validlapserandomcoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with Zero co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with Zero co-ordinates and valid lapse time"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "validlapsezerocoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with Negative  co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with Negative  co-ordinates and valid lapse time"
Environment("ExpectedResult") = "Error should be thrown"
blnResult =  VerifyLongClick(objMobiEdit , "validlapsenegativecoords") 

'Step : Execute LongClick  with only X  co-ordinate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with only X Co-ordinate and valid lapse time"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "validlapsexcoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with only Y  co-ordinate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with only Y  Co-ordinate and valid lapse time"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "validlapseycoords") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with valid  values 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with valid X , Y and Lapsetime value"
Environment("ExpectedResult") = "LongClick  should work correctly."
blnResult =  VerifyLongClick(objMobiEdit , "withvalidlapsevalidvalue") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'Step : Execute LongClick  with invalid lapse time
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick ." & VBNewLine
Environment("Description") = " Execute method with invalid lapse time"
Environment("ExpectedResult") = "Error  should be thrown"
blnResult =  VerifyLongClick(objMobiEdit , "withinvalidlapsetime") 

'Code to close keypad
strImagePathAfter = Environment("SystemTempDir") & "\" & Replace(Replace(Replace(CStr(Now), "/", ""), " ", "_"), ":", "") & ".bmp" 
MobiDevice("Phone Lookup").MobiButton("SignIn").CaptureBitmap  strImagePathAfter , True

If CompareBitmap(strImagePathAfter , strImagePathBefore)  Then
	MobiDevice("Phone Lookup").ButtonPress eBACK
	Wait 1
End If

'Put away focus from Username
MobiDevice("Phone Lookup").MobiEdit("ed_Password").Set "demo"
Wait 1

'#########################################################
'End test iteration
EndTestIteration()












