
'##########################################################################################################
' Objective: Launch  PhoneLookup app and test MobiCheckBox  Object
' Test Description: Execute all methods for MobiCheckBox on Remember Me Checkbox 
' Steps:
' Step1:Verify Login Screen
' Step2:Execute CaptureBitmap with .png file
' Step3:ExecuteCaptureBitmap with .bmp file
' Step4:Execute CaptureBitmap with .override.bmp file
' Step5:Execute CaptureBitmap with override .png file
' Step6:Execute CheckProperty when object is  visible
' Step7:Execute CheckProperty when object is not visible
' Step8:Execute ChildObjects
' Step9:Execute Click with  Boundary coordinates
' Step10:Execute Click with  Random coordinates
' Step11:Execute Click with  Zero coordinates
' Step12Execute Click without coordinates
' Step13Execute Click Negative coordinates
' Step14:Execute Click X coordinates
' Step15Execute Click Y coordinates
' Step16:Execute Click Valid X&y coordinates
' Step17:Execute Exist when object is visible
' Step18:Execute Exist when object is not visible
' Step19:Execute GetROProperty
' Step20:Execute GetTOProperties
' Step21:Execute GetToProperty
' Step22:Execute Refresh
' Step23:ExecuteToString
' Step24:Execute WaitProperty when object is visible 
' Step25:Execute WaitProperty when object is not visible 
' Step26:Execute Set checked
' Step27Execute Set checked
' Step28:Execute Set unchecked
' Step29:Execute Set unchecked
' Step30:Execute Set without any parameter
' Step31:Execute Set To Property
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
intStep = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""

'#######################################################

' Step1: Navigate to Search screen
'Expected Result: Search screen should be displayed
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Verify Login Screen" & VBNewLine
Environment("ExpectedResult") = "Login Screen should be displayed"

'Set object for Button
Set objMobiCheck = MobiDevice("Phone Lookup").MobiCheckbox("RememberMe")

'Call function to createreporttemplare
CreateReportTemplate()


'Call navigate to screen function 
strResult  = Cstr(NavigateScreenOnPhoneLookup("Login"  , objMobiCheck , ""))


' Step2:  Execute CaptureBitmap with .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on Mobicheckbox." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the png file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiCheck , "png")

' Step3:  Execute CaptureBitmap with .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute capture bitmap with .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on Mobicheckbox." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the bmp file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiCheck , "bmp")

' Step4:  Execute CaptureBitmap  with override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "excute capture bitmap with override .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on Mobicheckbox." & VBNewLine
Environment("ExpectedResult") = "Error message for override of bmp image should appear."
strResult = VerifyCaptureBitmap(objMobiCheck , "override_bmp")

' Step5:  Execute CaptureBitmap with override .png file 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute capture bitmap with override .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on Mobicheckbox." & VBNewLine
Environment("ExpectedResult") = " Error message for override of png image should appear."
strResult = VerifyCaptureBitmap(objMobiCheck , "override_png")


' Step 6:  Execute 'CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check propertywhen object is  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."

strResult = VerifyCheckProperty(objMobiCheck, "visible", True, 5000, True)


'Step 8 : Execute ChildObjects
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Child objects"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
strResult = VerifyChildObjects(objMobiCheck , "recursive" , 1)

'Step 8 : Execute ChildObjects
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Child objects"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
strResult = VerifyChildObjects(objMobiCheck , "nonrecursive" , 1)

'Step 9 : Execute Click with  Boundary coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiCheckBox with Boundary coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiCheckBox with Boundary coordinates." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiCheck, "withboundarycoords")

'Step 22:  Execute Click with boundary coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton with boundarycoords top left"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundarycoords top left"
blnStepRC = VerifyClick(objMobiCheck, "withboundarycoordstopleft")

'Step 22:  Execute Click with with boundarycoords top right coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton with boundarycoords top right"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundarycoords top right"
blnStepRC = VerifyClick(objMobiCheck, "withboundarycoordstopright")

'Step 22:  Execute Click with with boundarycoords bottom  left
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiButton with boundarycoords bottom  left"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with boundarycoords bottom  left"
blnStepRC = VerifyClick(objMobiCheck, "withboundarycoordsbottomleft")

'Step 10: Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")= "Execute Click on MobiCheckBox with Random coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiCheckBox with  Random coordinates." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult =  VerifyClick(objMobiCheck, "withrandomcoords")


'Step 12 : Execute Click withoutcoordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")= "Execute Click on MobiCheckBox withoutcoordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiCheckBox without coordinates." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult =  VerifyClick(objMobiCheck, "withoutcoords")



'Step 14: Execute Click only X coordinate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")= "Execute Click on MobiCheckBox with X coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiCheckBox X coordinates." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult =  VerifyClick(objMobiCheck, "withxvalue")

'Step 15 : Execute Click only Y coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")= "Execute Click on MobiCheckBox with Y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiCheckBox Y coordinates." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult =  VerifyClick(objMobiCheck, "withyvalue")

'Step 16: Execute Click at valid X&Y
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")= "Execute Click on MobiCheckBox with X &Y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiCheckBox Valid X & Y coordinates." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult =  VerifyClick(objMobiCheck, "withvalidvalue")


'Step 17: Execute Exist when object is  visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = VerifyExist(objMobiCheck, True, 5)


'Step 19 : Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetRoProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."

arrProps = Array("id", "text")
arrValues = Array("2131230796", "Remember Me")
strResult = VerifyGetROProperty(objMobiCheck,arrProps,arrValues)


'Step 20: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetToProperties"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrprop = Array ("text","micclass")

strResult = VerifyGetTOProperties(objMobiCheck,arrprop)


'Step 21 : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("id", "text")
arrValues = Array("2131230796", "Remember Me")
strResult = VerifyGetTOProperty(objMobiCheck, arrProps, arrValues)


'Step 22  Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Refresh object"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiCheck)

'Step 23 : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
strResult = VerifyTOString(objMobiCheck)


'Step 24 : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."

strResult = VerifyWaitProperty(objMobiCheck, "visible", True , 5000, True)

'Step 26: Execute 'Set checked
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set  checked"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Object should get checked"
objMobiCheck.Set eUNCHECKED
strResult = VerifySet(objMobiCheck,  eCHECKED , null)


'Step 27 Execute 'Set checked
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Checked"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Object should remain checked."

strResult = VerifySet(objMobiCheck, eCHECKED, null)


'Step 28: Execute 'Set  unchecked
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Unchecked"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Object should get un-checked"
strResult = VerifySet(objMobiCheck,  eUNCHECKED , null)


'Step 29: Execute 'Set unchecked
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Unchecked"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Object should remain un-checked"

strResult = VerifySet(objMobiCheck,  eUNCHECKED , null)


'Step 30: Execute 'Set without any parameter
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set without any parameter"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Error message should be thrown"

strResult = VerifySet(objMobiCheck,null,null)


'Step 31: Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiCheckbox." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
strResult = VerifySetTOProperty(objMobiCheck, arrProps)

'NavigateScreenOnPhoneLookup  "Controls" , MobiDevice("Phone Lookup").MobiDatetimePicker("DatePicker") , "DatePicker"
Login "mobilelabs" , "demo"
' Step 7:  Execute 'CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check propertywhen object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."

strResult = VerifyCheckProperty(objMobiCheck, "visible", True, 15000, False)
			
'Step 18: Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is not visible "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
'execute exist method
strResult = VerifyExist(objMobiCheck, False, 10)

'Step 25 : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."

'execute wait property
strResult = VerifyWaitProperty(objMobiCheck, "visible", True, 15000, False)

'******************************************************************************************************************************************************************

'Call function to end test iteration
EndTestIteration()


















