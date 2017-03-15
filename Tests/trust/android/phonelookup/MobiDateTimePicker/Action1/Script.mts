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
arrToPropValues = Array("True" , 1)
arrROProps = Array("name" , "nativeclass")
arrROPropValues = Array("DatetimePicker" , "android.widget.DatePicker")

'Create an html report template
CreateReportTemplate()

'*********************************************************************************************************************
' Step : Navigate to PickerView page of PhoneLookUp App.
'Expected Result: User should be navigated to PickerView Page

'Set object for DateTime Picker 
Set objMobiDateTime = MobiDevice("Phone Lookup").MobiDateTimePicker("DatePicker")

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Navigate to PickerView page of PhoneLookUp App." & VBNewLine
Environment("ExpectedResult") = "User should be navigated to PickerView Page"

'Call navigate to screen function 
NavigateScreenOnPhoneLookup "Controls"  , objMobiDateTime , "DatePicker" 

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


' Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiDateTimePicker recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiDateTime,"recursive",0)

' Step 7:  Execute  ChildObjects nonrecusrively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiDateTimePicker non-recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiDateTime,"nonrecursive",0)

' Step 13:  Execute Click  with boundary coordinates at Top-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiDatePicker."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime, "withboundarycoordsTopLeft")
'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
End If
' Step 13:  Execute Click  with boundary coordinates at Top-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Right corner on MobiDatePicker."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime, "withboundarycoordsTopRight")
'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
End If
' Step 13:  Execute Click  with boundary coordinates at Bottom-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Left corner on MobiDatePicker."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime, "withboundarycoordsBottomLeft")
'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
End If
' Step 13:  Execute Click  with boundary coordinates at Bottom-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Right corner on MobiDatePicker."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDateTime, "withboundarycoordsBottomRight")
'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 2
End If

'Step  : Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiButton with Random coordinates." & VBNewLine
Environment("Description") = "Click : Execute method with Random co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiDateTime, "withrandomcoords")


'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
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


'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
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

'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
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

'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
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

'Return back to Object screen
If  blnResult   AND objMobiDateTime.Exist(10) Then
	MobiDevice("micclass:=MobiDevice").ButtonPress eBACK
	Wait 5
End If

If Not objMobiDateTime.Exist(4)  Then
	MobiDevice("Phone Lookup").MobiButton("name:=Change.*").Click
	Wait 4
End If

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


'Step  : Execute Select with only valid year
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select with only valid year" & VBNewLine
Environment("Description")  = "Select : Execute method to select valid year only"
Environment("ExpectedResult") = "Year should get selected"
strItem = "2014, , , ,"
strResult =  VerifySelect(objMobiDateTime , "IntegerInput" , strItem , null)



'Step  : Execute Select  with valid  Year and Month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select  with valid  Year and Month"  & VBNewLine
Environment("Description")  = "Select : Execute method to select valid year and month only"
Environment("ExpectedResult") = "Year and Month should get selected"
strItem = "2014,10 , , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)

'Step : Execute Select  with valid  Year and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select  with valid  Year and  Day"  & VBNewLine
Environment("Description")  = "Select : Execute method to select valid year and Day only"
Environment("ExpectedResult") = "Year and Day should get selected"
strItem = "2014, , 20, ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)

'Step  : Execute Select  with valid  Month  and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select  with valid  Year and Month"  & VBNewLine
Environment("Description")  = "Select : Execute method to select valid Month  and Day only"
Environment("ExpectedResult") = "Month  and Day should get selected"
strItem = ",10 , 20, ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)



'Step : Execute Select  only  valid  month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute Select with valid month"
Environment("ExpectedResult") = "Month should get selected"
strItem = " ,10 , , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)

'Step  : Execute Select  only  Invalid  month
'##########################################################

'Step : Execute Select  with valid Year , month and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select  with valid Year , month and  Day" & VBNewLine
Environment("Description")  = "Select : Execute Select on  valid year , month and day"
Environment("ExpectedResult") = "Valid Year , month and Day should get selected"
strItem = "2004 ,4 ,27 , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)



'Step  : Execute Select  on a leap year and 29 days
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select  on a leap year and 29 days" & VBNewLine
Environment("Description")  = "Select : Execute method for selecting 29 days on a leap year"
Environment("ExpectedResult") = "29 days should get selected with a leap year"
strItem = "2000 ,2 ,29 , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)


'Step  : Execute Select  with only valid Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select  with only valid Day" & VBNewLine
Environment("Description")  = "Select : Execute method to select a valid day"
Environment("ExpectedResult") = "Day should get selected"
strItem = " , ,20 , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)


'Step  : Execute Select  year which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute method to select an already selected Year"
Environment("ExpectedResult") = "Year value should not change"

'Select  an year , month and day before execution
objMobiDateTime.Select  2014 , 2, 27
wait 3
strItem = "2014 , , , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)

'Step : Execute Select  Month  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute method to select an already selected month"
Environment("ExpectedResult") = "Month value should not change"

strItem = " , 2, , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)

'Step  : Execute Select  Day  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute method to select an already selected Day"
Environment("ExpectedResult") = "Day's value should not change"

strItem = " , ,27 , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)

'Step  : Execute Select  year , month and day which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intSubStep & ": " &_
"Execute Select on MobiButton." & VBNewLine
Environment("Description")  = "Select : Execute method to select an already selected year , month and day"
Environment("ExpectedResult") = "No change in selection should occur"

strItem = "2014 ,2 ,27 , ,"
strResult = VerifySelect(objMobiDateTime, "IntegerInput" , strItem , null)


'Select method with string as input


'Step  : Execute Select with only valid year
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select valid year only"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "2014"
strResult =  VerifySelect(objMobiDateTime , "stringinput" , strItem , null)



'Step  : Execute Select  with valid  Year and Month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select valid year and month only"
Environment("ExpectedResult") = "Year and Month should get selected"
strItem = "2014-10"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step : Execute Select  only  valid  month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute Select with valid month"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "10"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)


'Step : Execute Select  with valid Year , month and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute Select on  valid year , month and day"
Environment("ExpectedResult") = "Valid Year , month and Day should get selected"
strItem = "2012-4-22"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step  : Execute Select  on a leap year and 29 days
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method for selecting 29 days on a leap year"
Environment("ExpectedResult") = "29 days should get selected with a leap year"
strItem = "2000-2-29"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)


'Step  : Execute Select  with only valid Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select a valid day"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "40"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step  : Execute Select  year which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select an already selected Year"
Environment("ExpectedResult") = "Year value should not change"

'Select  an year , month and day before execution
objMobiDateTime.Select  "2014-2-2"
wait 3
strItem = "2014-2-2"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step : Execute Select  Month  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select an already selected month"
Environment("ExpectedResult") = "Month value should not change"

strItem = "2014-2"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step  : Execute Select  Day  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select an already selected Day"
Environment("ExpectedResult") = "Day's value should not change"

strItem = " 2-2"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step  : Execute Select  year , month and day which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

Environment("Description")  = " Execute method to select an already selected year , month and day"
Environment("ExpectedResult") = "No change in selection should occur"

strItem = "2014-2-2"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

'Step  : Execute Select  both Date and Time
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

Environment("Description")  = "Select both Date and Time"
Environment("ExpectedResult") = "Date and Time should get selected"

strItem = "1999-4-4 t 10:10:10"
strResult = VerifySelect(objMobiDateTime, "stringinput" , strItem , null)

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

'Navigate to Login screen 
LogOut 

' Step :  Execute  CheckProperty when object is not  visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute  CheckProperty when object is not visible" & VBNewLine
Environment("Description") = "CheckProperty : Execute method to check property value when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult =  VerifyCheckProperty(objMobiDateTime, "datepickermode", 1 , 5000, False)

'Step  Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiButton when object is visible" & VBNewLine
Environment("Description") = "Exist : Verify method when object is not visible"
Environment("ExpectedResult") = "Exist should return False when object is not visible."
blnResult = VerifyExist(objMobiDateTime, False, 10)

'Step : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty when object is not visible on MobiButton." & VBNewLine
Environment("Description")  =  "WaitTOProperty : Execute method to wait till  property attains value"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and return False."
strResult =  VerifyWaitProperty(objMobiDateTime, "visible", True , 5000, False)

'*********************************************************************************************************************
'*********************************************************************************************************************
EndTestIteration


