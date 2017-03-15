'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiDatePicker & MobiTimePicker object
' Test Description: Execute all methods for MobiDatePicker & MobiTimePicker  in Controls of PhoneLookup app
'#########################################################################################################

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

'Input values
arrTOProps = Array("visible"  , "datepickermode")
arrToPropValues = Array("True" , 0)
arrROProps = Array("name" , "nativeclass")
arrROPropValues = Array("DatetimePicker" , "android.widget.TimePicker")

'Create an html report template
CreateReportTemplate()

'*********************************************************************************************************************
' Step  : Navigate to TimePicker Screen of PhoneLookUp App. 
'Expected Result: User should be navigated to TimePicker Screen

'Set object for TimePicker
Set objMobiDateTime = MobiDevice("Phone Lookup").MobiDateTimePicker("TimePicker")

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Navigate to TimePicker Screen of PhoneLookUp App." & VBNewLine
Environment("Description")  = "Open Time Picker screen"
Environment("ExpectedResult") = "User should be navigated to TimePicker Screen"

'Call navigate to screen function 
strResult  = NavigateScreenOnPhoneLookup("Controls"  , objMobiDateTime , "TimePicker")

'*********************************************************************************************************************
' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap with .png format" & VBNewLine
Environment("Description") = "CaptureBitMap : Execute method to capture image in .png format"
Environment("ExpectedResult") = "Image should get captured in .png format"
blnResult = VerifyCaptureBitmap(objMobiDateTime , "png")

' Step:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap with .bmp format" & VBNewLine
Environment("Description") = "CaptureBitMap : Execute method to capture image in .bmp format"
Environment("ExpectedResult") = "Image should get captured in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiDateTime , "bmp")

'Step :  Execute CaptureBitmap to override an .bmp image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap  to override an .bmp image" & VBNewLine
Environment("Description") = "CaptureBitMap : Execute method  to override an .bmp image"
Environment("ExpectedResult") = "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiDateTime , "override_bmp")

'Step :  Execute CaptureBitmap to override an .png image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap  to override an .png image" & VBNewLine
Environment("Description") = "CaptureBitMap : Execute method  to override an .png image"
Environment("ExpectedResult") =  "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiDateTime , "override_png")

'Step :  Execute  CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute  CheckProperty when object is visible" & VBNewLine
Environment("Description") = "CheckProperty : Execute method to check property value when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiDateTime, "datepickermode", 1 , 5000, True)


'Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute childobject method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on mobilist." & VBNewLine

Environment("ExpectedResult") = "Return child object recursively in the application"
'blnFlag = VerifyChildObjects(objMobiDateTime  ,"recursive",26)
blnFlag = VerifyChildObjects(objMobiDateTime, "recursive" , 0)

 

' Step 7:  Execute  ChildObjects nonrecusrively
'#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "ChildObjects : Execute ChildObjects on MobiTimePicker non-recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiSlider." & VBNewLine
'Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
'blnFlag = VerifyChildObjects(objMobiDateTime  ,"nonrecursive",0)
'
' 'Step 13:  Execute Click  with boundary coordinates at Top-Left corner
''#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime ,"withboundarycoordsTopLeft")

'
'' Step 13:  Execute Click  with boundary coordinates at Top-Right corner
''#######################################################
'
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Right corner on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime,"withboundarycoordsTopRight")

'
'' Step 13:  Execute Click  with boundary coordinates at Bottom-Left corner
''#######################################################
'
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Left corner on MobiList"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime,"withboundarycoordsBottomLeft")

'
' Step 13:  Execute Click  with boundary coordinates at Bottom-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Right corner on MobiDropdown."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime,"withboundarycoordsBottomRight")




'Step  Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiButton when object is visible" & VBNewLine
Environment("Description") = "Exist : Verify method when object is visible"
Environment("ExpectedResult") = "Exist should return True when object is visible"
blnResult = VerifyExist(objMobiDateTime, True, 5)


'Step : Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiButton." & VBNewLine
Environment("Description") = "GetROProperty : Verify object run time values"
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnResult = VerifyGetROProperty(objMobiDateTime , arrROProps , arrROPropValues)

'Step  : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiButton." & VBNewLine
Environment("Description") = "GetTOProperties : Verify object description properties"
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
blnResult = VerifyGetTOProperties(objMobiDateTime, arrTOProps )

'Step : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiButton." & VBNewLine
Environment("Description") = "GetTOProperties : Verify object description propertie and their values"
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
blnResult =  VerifyGetTOProperty(objMobiDateTime, arrTOProps , arrTOPropValues)

'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiButton." & VBNewLine
Environment("Description")  = "RefreshObject : Re-identify object in the application"
Environment("ExpectedResult") = "Object should get re - identified"
strResult =  VerifyRefreshObject(objMobiDateTime)


'Step : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute method to fetch object type an class"
Environment("ExpectedResult") = "TOString should return the object type and class."
strResult =  VerifyTOString(objMobiDateTime)


'Step : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty when object is visible on MobiButton." & VBNewLine
Environment("Description")  = "WaitTOProperty : Execute method to wait till  property attains value"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and should return True"
strResult = VerifyWaitProperty(objMobiDateTime, "visible", True, 5000, True)


'Step   : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute method to set  property value"
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."

strResult =  VerifySetTOProperty(objMobiDateTime, arrTOProps)
'*********************************************************************************************************************

'Step : Execute Select  with valid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour value"
Environment("ExpectedResult") = "Hour should get selected"
strValue = " , , , 8 , "
blnResult = VerifySelect(objMobiDateTime , "IntegerInput" , strValue , Null)



'Step : Execute Select  with valid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour value"
Environment("ExpectedResult") = "Hour should get selected"
strValue = " , , , 23 , "
blnResult = VerifySelect(objMobiDateTime , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour value"
Environment("ExpectedResult") = "Hour should get selected"
strValue = " , , , 1 , "
blnResult = VerifySelect(objMobiDateTime , "IntegerInput" , strValue , Null)



'Step : Execute Select  with valid  Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid minute value"
Environment("ExpectedResult") = "Minute should get selected"
strValue = " , , ,  , 59"
blnResult = VerifySelect(objMobiDateTime , "IntegerInput" , strValue , Null)



'Step : Execute Select  with valid  hour and  Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select withvalid hour and minute value"
Environment("ExpectedResult") = "Specified hour and minute value should get selected"
strValue = " , , , 10 , 45"
blnResult = VerifySelect(objMobiDateTime , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  hour and  Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select withvalid hour and minute value"
Environment("ExpectedResult") = "Specified hour and minute value should get selected"
strValue = " , , , 23 , 59"
blnResult = VerifySelect(objMobiDateTime , "IntegerInput" , strValue , Null)



'Step : Execute Select  with both date and time.
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with both Date and Time in valid format"
Environment("ExpectedResult") = "Time should get selected"
strValue = "2013-10-9 t 11:23:20"
blnResult = VerifySelect(objMobiDateTime , "stringinput" , strValue , Null)




''Step  Execute Exist when object is visible
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify method when object is visible"
'Environment("ExpectedResult") = "Exist should return True when object is visible"
'blnResult = VerifyExist(objMobiDateTime, True, 5)




'Step  : Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton with Random coordinates." & VBNewLine
Environment("Description") = "Click : Execute method with Random co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiDateTime, "withrandomcoords")


If  blnResult  AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(10)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 5
End If



'Step : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton without coordinates." & VBNewLine
Environment("Description") = "Click : Execute method without  co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTime, "withoutcoords")


If  blnResult  AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(10)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 5
End If




'Step  : Execute Click  at only one co-ordinate (Only X)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton with only X coordinate" & VBNewLine
Environment("Description") = "Click : Execute method with only X co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTime, "withxvalue")

If  blnResult  AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(10)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 5
End If

'Step  : Execute Click  at only one co-ordinate (Only Y)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton with only Y coordinate" & VBNewLine
Environment("Description") = "Click : Execute method with only Y co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTime, "withyvalue")

If  blnResult  AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(10)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 5
End If

'Step  : Execute Click  at  any valid value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click  at  any valid value." & VBNewLine
Environment("Description") = "Click : Execute method with any valid co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTime, "withoutcoords")

If  blnResult  AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(10)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 5
End If


'Step   : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify waitproperty when object is visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value"
blnResult = VerifyWaitProperty(objMobiDateTime, "visible", True , 5000, True)

LogOut

'Step  Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify method when object is not visible"
Environment("ExpectedResult") = "Exist should return False when object is not visible."
blnResult = VerifyExist(objMobiDateTime, False, 5)

'Step  : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify waitproperty when object is not visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and return False"
blnResult = VerifyWaitProperty(objMobiDateTime, "visible",True, 5000, False)


'Step :  Execute  CheckProperty when object is  not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute  CheckProperty when object is  not visible" & VBNewLine
Environment("Description") = "CheckProperty : Execute method to check property value when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiDateTime, "datepickermode", 1 , 5000 , False)

'*********************************************************************************************************************
'End test iteration
EndTestIteration()


