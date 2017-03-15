


'##########################################################################################################
'Objective: Login to UICatalog and Test DateTimePicker with mode 1
' Test Description: Execute all MobiDateTimePicker methods
'##########################################################################################################

'#######################################################
'Declare Variables
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
'#######################################################

'#######################################################
'Initializations
intStep = 0
Environment("Component") = "UICatalog_ObjectBased"
Environment("MethodName")  = ""
Environment("intStepNo") = 0
Environment("Status") = ""

'#######################################################

'Input values
arrTOProps = Array("visible" ,  "datepickermode")
arrToPropValues = Array(True , 1)
arrROProps = Array("name" , "nativeclass")
arrROPropValues = Array("DatetimePicker" , "UIDatePicker")

'Create an html report template
CreateReportTemplate()

'#######################################################
' Step: Navigate to UIPicker Screen
'Expected Result: UIPicker screen should be displayed
Environment("StepName") = "Step" & intStep
Environment("ExpectedResult") = "UIPicker  screen should be displayed"

'Set object for Button
Set objMobiDateTimePicker = MobiDevice("UICatalog").MobiDatetimePicker("DatetimePicker_Mode2")

'Call function to navigate to UIPicker screen
blnFlag = NavigateToObjectScreenUICatalog (objMobiDateTimePicker  , 2 , "datepicker"  , "pickers")
If Not blnFlag Then
	ReportStep "SelectDatePicker" , "Screen should be displayed with UIDatePicker Mode 1 object on it" , "Failed to open" , "N/A"
	EndTestIteration()
End If 

'###########################################################

' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method to capture image in .png format"
Environment("ExpectedResult") = "Image should get captured in .png format" 
blnResult = VerifyCaptureBitmap(objMobiDateTimePicker , "png")

' Step:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method to capture image in .bmp format"
Environment("ExpectedResult") = "Image should get captured in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiDateTimePicker , "bmp")

' Step :  Execute CaptureBitmap to override an .bmp image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method  to override an .bmp image"
Environment("ExpectedResult") = "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiDateTimePicker , "override_bmp")

' Step :  Execute CaptureBitmap to override an .png image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method  to override an .png image"
Environment("ExpectedResult") =  "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiDateTimePicker , "override_png")

' Step :  Execute  CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method to check property value when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiDateTimePicker, "visible" , True , 5000 , True)

'Navigate to other screen
MobiDevice("UICatalog").MobiButton("btnBack").Click
Wait 3

' Step :  Execute  CheckProperty when object is not  visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method to check property value when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiDateTimePicker, "visible" , True , 5000 , False)

'Navigate back to object screen
NavigateToObjectScreenUICatalog objMobiDateTimePicker  , 2 , "datepicker"  , "pickers"

'Step  : Execute ChildObjects
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verfiy child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) else 0"
blnResult = VerifyChildObjects(objMobiDateTimePicker)

'Step  : Execute Click with  Boundary coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with boundary co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiDateTimePicker, "withboundarycoords")


'Step  : Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with Random co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiDateTimePicker, "withrandomcoords")


'Step : Execute Click with  Zero coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with zero co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withzerovalues")
NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 2 , "datepicker"  , "pickers" 

'Step  : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method without  co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withoutcoords")


'Step  : Execute Click  at negative co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with negative co-ordinates"
Environment("ExpectedResult") = "Click should throw error message"
blnResult =  VerifyClick(objMobiDateTimePicker, "withnegativecoords")


'Step  : Execute Click  at only one co-ordinate (Only X)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with only X co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withxvalue")


'Step  : Execute Click  at only one co-ordinate (Only Y)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with only Y co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withyvalue")

'Step  : Execute Click  at  any valid value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with any valid co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withvalidvalue")


'Step  Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method when object is visible"
Environment("ExpectedResult") = "Exist should return True when object is visible"
blnResult = VerifyExist(objMobiDateTimePicker, True, 5)

'Navigate to other screen
MobiDevice("UICatalog").MobiButton("btnBack").Click
Wait 3

'Step  Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify method when object is not visible"
Environment("ExpectedResult") = "Exist should return False when object is not visible."
blnResult = VerifyExist(objMobiDateTimePicker, False, 5)

'Navigate back to object screen
NavigateToObjectScreenUICatalog objMobiDateTimePicker  , 2 , "datepicker"  , "pickers"

'Step  Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object run time values"
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."

blnResult = VerifyGetROProperty(objMobiDateTimePicker , arrROProps , arrROPropValues)

'Step  : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object description properties"
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
blnResult = VerifyGetTOProperties(objMobiDateTimePicker, arrTOProps)

'Step : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify object description propertie and their values"
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
blnResult =  VerifyGetTOProperty(objMobiDateTimePicker, arrTOProps, arrToPropValues)


'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object refresh"
Environment("ExpectedResult") = "RefreshObject should re-identify  the object in the application"
blnResult = VerifyRefreshObject(objMobiDateTimePicker)


'Step  : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object  type and class"
Environment("ExpectedResult") = "ToString should return the object type and class."
blnResult = VerifyTOString(objMobiDateTimePicker)

'Step   : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify waitproperty when object is visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value"
blnResult = VerifyWaitProperty(objMobiDateTimePicker, "visible", True , 5000, True)

'Navigate to other screen
MobiDevice("UICatalog").MobiButton("btnBack").Click
Wait 3

'Step  : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify waitproperty when object is not visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and return False"
blnResult = VerifyWaitProperty(objMobiDateTimePicker, "visible",True, 5000, False)

'Navigate back to object screen
NavigateToObjectScreenUICatalog objMobiDateTimePicker  , 2 , "datepicker"  , "pickers"

'Step  : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify property values after update"
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnResult = VerifySetTOProperty(objMobiDateTimePicker, arrTOProps)

'Step  : Execute Select with only valid year
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select valid year only"
Environment("ExpectedResult") = "Year should get selected"
strItem = "2014, , , ,"
strResult =  VerifySelect(objMobiDateTimePicker , "IntegerInput" , strItem , null)


'Step  : Execute Select  with only  Invalid year
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select invalid year only"
Environment("ExpectedResult") = "Error message should get thrown"
strItem =  "200000 , , , , " 
strResult =  VerifySelect(objMobiDateTimePicker , "IntegerInput" , strItem , null)

'Step  : Execute Select  with valid  Year and Month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select valid year and month only"
Environment("ExpectedResult") = "Year and Month should get selected"
strItem = "2014,10 , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step : Execute Select  with valid  Year and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select valid year and Day only"
Environment("ExpectedResult") = "Year and Day should get selected"
strItem = "2014, , 20, ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  with valid  Month  and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select valid Month  and Day only"
Environment("ExpectedResult") = "Month  and Day should get selected"
strItem = ",10 , 20, ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)


'Step : Execute Select  with Invalid Year  as Zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute Select  with Invalid Year  as Zero"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "0 , , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step : Execute Select  with Invalid Month  as Zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select :Execute Select  with Invalid Month  as Zero"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = " ,0 , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  with Invalid  Day  as Zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute Select  with Invalid  Day  as Zero"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = " , ,0 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step : Execute Select  only  valid  month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute Select with valid month"
Environment("ExpectedResult") = "Month should get selected"
strItem = " ,10 , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  only  Invalid  month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute Select with Invalid month"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = " ,40 , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step : Execute Select  with valid Year , month and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute Select on  valid year , month and day"
Environment("ExpectedResult") = "Valid Year , month and Day should get selected"
strItem = "2004 ,4 ,27 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)


'Step : Execute Select  only  Invalid  Year as negative value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method for selecting invalid year having  negative value"
Environment("ExpectedResult") = "Error message should be thrown."
strItem = "-20 , , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)


'Step  : Execute Select  only  Invalid  Month as negative value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method for selecting invalid month having  negative value"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = " ,-20 , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  only  Invalid  Day as negative value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method for selecting invalid day having negative value"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = " , ,-20 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  on a leap year and 29 days
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method for selecting 29 days on a leap year"
Environment("ExpectedResult") = "29 days should get selected with a leap year"
strItem = "2000 ,2 ,29 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select any non leap year  day greater than 29 days in feb
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method for selecting an non leap year and 29 days"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "2013 ,2 ,29 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)


'Step  : Execute Select  with only valid Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select a valid day"
Environment("ExpectedResult") = "Day should get selected"
strItem = " , ,20 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  with only Invalid Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select a  Invalid day"
Environment("ExpectedResult") = "Error message should be thrown."
strItem = " , ,40 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  year which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select an already selected Year"
Environment("ExpectedResult") = "Year value should not change"

'Select  an year , month and day before execution
objMobiDateTimePicker.Select  2014 , 2, 27

strItem = "2014 , , , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step : Execute Select  Month  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select an already selected month"
Environment("ExpectedResult") = "Month value should not change"

strItem = " , 2, , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  Day  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select an already selected Day"
Environment("ExpectedResult") = "Day's value should not change"

strItem = " , ,27 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)

'Step  : Execute Select  year , month and day which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

Environment("Description")  = " Execute method to select an already selected year , month and day"
Environment("ExpectedResult") = "No change in selection should occur"

strItem = "2014 ,2 ,27 , ,"
strResult = VerifySelect(objMobiDateTimePicker, "IntegerInput" , strItem , null)


'Select  new scenarios with string input

'Step  : Execute Select with only valid year
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select valid year only"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "2014"
strResult =  VerifySelect(objMobiDateTimePicker , "stringinput" , strItem , null)


'Step  : Execute Select with only invalid year and valid month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select  with invalid year and valid month"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "30000-10"
strResult =  VerifySelect(objMobiDateTimePicker , "stringinput" , strItem , null)

'Step  : Execute Select  with valid  Year and Month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select valid year and month only"
Environment("ExpectedResult") = "Year and Month should get selected"
strItem = "2014-10"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  with valid  Month  and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select valid Month  and Day only"
Environment("ExpectedResult") = "Month  and Day should get selected"
strItem = "9-30"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  with invalid  Month  and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select invalid Month  and  valid Day"
Environment("ExpectedResult") = "Month  and Day should get selected"
strItem = "13-10"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step : Execute Select  with Invalid Year  as Zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute Select  with Invalid Year  as Zero"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "0000-2"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step : Execute Select  with Invalid Month  as Zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute Select  with Invalid Month  as Zero"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "2013-00"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  with Invalid  Day  as Zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute Select  with Invalid  Day  as Zero"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "2013-4-0"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step : Execute Select  only  valid  month
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute Select with valid month"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "10"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)


'Step : Execute Select  with valid Year , month and  Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute Select on  valid year , month and day"
Environment("ExpectedResult") = "Valid Year , month and Day should get selected"
strItem = "2012-4-22"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  on a leap year and 29 days
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method for selecting 29 days on a leap year"
Environment("ExpectedResult") = "29 days should get selected with a leap year"
strItem = "2000-2-29"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select any non leap year  day greater than 29 days in feb
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method for selecting an non leap year and 29 days"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "2013-2-29"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)


'Step  : Execute Select  with only valid Day
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select a valid day"
Environment("ExpectedResult") = "Error message should be thrown"
strItem = "40"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  year which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = " Execute method to select an already selected Year"
Environment("ExpectedResult") = "Year value should not change"

'Select  an year , month and day before execution
objMobiDateTimePicker.Select  "2014-2-2"

strItem = "2014-2-2"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step : Execute Select  Month  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute method to select an already selected month"
Environment("ExpectedResult") = "Month value should not change"

strItem = "2014-2"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  Day  which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Select : Execute method to select an already selected Day"
Environment("ExpectedResult") = "Day's value should not change"

strItem = " 2-2"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  year , month and day which is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

Environment("Description")  = " Execute method to select an already selected year , month and day"
Environment("ExpectedResult") = "No change in selection should occur"

strItem = "2014-2-2"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)

'Step  : Execute Select  both Date and Time
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

Environment("Description")  = "Select both Date and Time"
Environment("ExpectedResult") = "Date and Time should get selected"

strItem = "1999-2-20 t 10:10:10"
strResult = VerifySelect(objMobiDateTimePicker, "stringinput" , strItem , null)


'*********************************************************************************************************************

'End test iteration
EndTestIteration()







