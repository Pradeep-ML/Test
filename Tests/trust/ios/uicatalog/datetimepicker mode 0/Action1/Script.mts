

'##########################################################################################################
'Objective: Login to UICatalog and Test DateTimePicker with mode 0
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
arrToPropValues = Array(True , 0)
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
Set objMobiDateTimePicker = MobiDevice("UICatalog").MobiDatetimePicker("DatetimePicker_Mode1")


'Call function to navigate to UIPicker screen
blnFlag = NavigateToObjectScreenUICatalog (objMobiDateTimePicker  , 1 , "datepicker"  , "pickers")
If Not blnFlag Then
	ReportStep "SelectDatePicker" , "Screen should be displayed with UIDatePicker Mode 0 object on it" , "Failed to open" , "N/A"
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
NavigateToObjectScreenUICatalog objMobiDateTimePicker  , 1 , "datepicker"  , "pickers"

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
NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

'Step  : Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with Random co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiDateTimePicker, "withrandomcoords")
NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

'Step : Execute Click with  Zero coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with zero co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withzerovalues")

NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

'Step  : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method without  co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withoutcoords")

NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

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

NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

'Step  : Execute Click  at only one co-ordinate (Only Y)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with only Y co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withyvalue")

NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

'Step  : Execute Click  at  any valid value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with any valid co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiDateTimePicker, "withvalidvalue")

NavigateToObjectScreenUICatalog  objMobiDateTimePicker  , 1 , "datepicker"  , "pickers" 

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


'Step  : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify property values after update"
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnResult = VerifySetTOProperty(objMobiDateTimePicker, arrTOProps)

'Step : Execute Select  with valid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour value"
Environment("ExpectedResult") = "Hour should get selected"
strValue = " , , , 8 , "
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour value"
Environment("ExpectedResult") = "Hour should get selected"
strValue = " , , , 1 , "
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour value"
Environment("ExpectedResult") = "Hour should get selected"
strValue = " , , , 12 , "
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with Invalid  hour value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with invalid hour value"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = " , , , 80 , "
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid minute value"
Environment("ExpectedResult") = "Minute should get selected"
strValue = " , , ,  , 59"
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with Invalid  Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with invalid minute value"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = " , , ,  , 99"
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  hour and  Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select withvalid hour and minute value"
Environment("ExpectedResult") = "Specified hour and minute value should get selected"
strValue = " , , , 10 , 45"
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with valid  hour and  invalid Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour and  invalid minute value"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = " , , , 10 , 145"
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)

'Step : Execute Select  with invalid  hour and  valid Minute value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with  invalid hour and  valid minute value"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = " , , , 100 , 45"
blnResult = VerifySelect(objMobiDateTimePicker , "IntegerInput" , strValue , Null)


'Select with string input

'Step : Execute Select  with valid hour 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = "10"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)


'Step : Execute Select  with valid  Minute 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid minute"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = "40"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)


'Step : Execute Select  with valid  hour and minute 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour and minute"
Environment("ExpectedResult") = "Hour and minute value should get selected"
strValue = "11:59"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)

'Step : Execute Select  with invalid  hour and valid minute
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with invalid hour and valid minute"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = "25:22"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)


'Step : Execute Select  with valid  hour and  invalid minute
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour and  invalid minute"
Environment("ExpectedResult") = "Error message should be thrown"
strValue = "09:61"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)


'Step : Execute Select  with valid  hour , minutes and seconds
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with valid hour , minutes and seconds"
Environment("ExpectedResult") = "Hour and minute should get selected"
strValue = "09:06:20"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)


'Step : Execute Select  with both date and time.
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with both Date and Time in valid format"
Environment("ExpectedResult") = "Time should get selected"
strValue = "2013-10-9 t 11:23:20"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)

'Step : Execute Select  with both date and time.
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with both Date and Time in invalid format."
Environment("ExpectedResult") = "Error messgae should be thrown"
strValue = "2013-10-9 t 01:03:02"
blnResult = VerifySelect(objMobiDateTimePicker , "stringinput" , strValue , Null)


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
NavigateToObjectScreenUICatalog objMobiDateTimePicker  , 1 , "datepicker"  , "pickers"

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
'*********************************************************************************************************************
'End test iteration
EndTestIteration()













