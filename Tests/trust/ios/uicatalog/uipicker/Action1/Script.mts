'##########################################################################################################
'Objective: Login to UICatalog and Test  Picker with two wheels
' Test Description: Execute all MobiPicker methods
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
Environment("WheelNumber") = ""
'#######################################################

'Input values
arrTOProps = Array("wheelcount" , "visible")
arrToPropValues = Array( 2 , True)

arrROProps = Array("nativeclass" ,  "name" , "itemscount")
arrROPropValues = Array("UIPickerView" , "Picker" , 7 )


'Set  Scroll/Swipe Objects
Set objTopWheelZero = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerTop_WheelZero")
Set objBottomWheelZero = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerBottom_WheelZero")
Set objBottomWheelOne = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerBottom_WheelOne")
Set objTopWheelOne  = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerTop_WheelOne")

'Create an html report template
CreateReportTemplate()

'#######################################################
' Step: Navigate to UIPicker Screen
'Expected Result: UIPicker screen should be displayed
Environment("StepName") = "Step" & intStep
Environment("ExpectedResult") = "UIPicker  screen should be displayed"

'Set object for Button
Set objMobiPicker = MobiDevice("UICatalog").MobiPicker("UIPicker")

'Call function to navigate to UIPicker screen
blnFlag = NavigateToObjectScreenUICatalog (objMobiPicker  ,  , "pickers" , "Pickers")
If Not blnFlag Then
	ReportStep "SelectPicker" , "Screen should be displayed with UIPicker object on it" , "Failed to open" , "N/A"
	EndTestIteration()
End If 

'###########################################################

' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method to capture image in .png format"
Environment("ExpectedResult") = "Image should get captured in .png format"
blnResult = VerifyCaptureBitmap(objMobiPicker , "png")

' Step:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method to capture image in .bmp format"
Environment("ExpectedResult") = "Image should get captured in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiPicker , "bmp")

' Step :  Execute CaptureBitmap to override an .bmp image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method  to override an .bmp image"
Environment("ExpectedResult") = "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiPicker , "override_bmp")

' Step :  Execute CaptureBitmap to override an .png image
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method  to override an .png image"
Environment("ExpectedResult") =  "Override error message should be thrown"
blnResult = VerifyCaptureBitmap(objMobiPicker , "override_png")

' Step :  Execute  CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method to check property value when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiPicker, "visible" , True , 5000 , True)

'Navigate to other screen
MobiDevice("UICatalog").MobiButton("btnBack").Click
Wait 3

' Step :  Execute  CheckProperty when object is not  visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method to check property value when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value."
blnResult = VerifyCheckProperty(objMobiPicker, "visible" , True , 5000 , False)

'Navigate back to object screen
NavigateToObjectScreenUICatalog  objMobiPicker ,  , "pickers" , "Pickers"

'Step  : Execute ChildObjects
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verfiy child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) else 0"
blnResult = VerifyChildObjects(objMobiPicker)

'Step  : Execute Click with  Boundary coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with boundary co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiPicker, "withboundarycoords")


'Step  : Execute Click with  Random coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with Random co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult = VerifyClick(objMobiPicker, "withrandomcoords")


'Step : Execute Click with  Zero coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with zero co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiPicker, "withzerovalues")
NavigateToObjectScreenUICatalog  objMobiPicker  ,  , "pickers" , "Pickers"

'Step  : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method without  co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiPicker, "withoutcoords")
NavigateToObjectScreenUICatalog  objMobiPicker  ,  , "pickers" , "Pickers"

'Step  : Execute Click  at negative co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with negative co-ordinates"
Environment("ExpectedResult") = "Click should throw error message"
blnResult =  VerifyClick(objMobiPicker, "withnegativecoords")


'Step  : Execute Click  at only one co-ordinate (Only X)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with only X co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiPicker, "withxvalue")


'Step  : Execute Click  at only one co-ordinate (Only Y)
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with only Y co-ordinate"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiPicker, "withyvalue")

'Step  : Execute Click  at  any valid value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute method with any valid co-ordinates"
Environment("ExpectedResult") = "Click should work correctly."
blnResult =  VerifyClick(objMobiPicker, "withvalidvalue")


'Step  Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method when object is visible"
Environment("ExpectedResult") = "Exist should return True when object is visible"
blnResult = VerifyExist(objMobiPicker, True, 5)

'Navigate to other screen
MobiDevice("UICatalog").MobiButton("btnBack").Click
Wait 2

'Step  Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify method when object is not visible"
Environment("ExpectedResult") = "Exist should return False when object is not visible."
blnResult = VerifyExist(objMobiPicker, False, 5)

'Navigate back to object screen
NavigateToObjectScreenUICatalog  objMobiPicker ,  , "pickers" , "Pickers"

'Step  Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object run time values"
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnResult = VerifyGetROProperty(objMobiPicker , arrROProps , arrROPropValues)

'Step  : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object description properties"
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
blnResult = VerifyGetTOProperties(objMobiPicker, arrTOProps)

'Step : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify object description propertie and their values"
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
blnResult =  VerifyGetTOProperty(objMobiPicker, arrTOProps, arrToPropValues)

'Step  : Execute GetVisibleText method with coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method with co-ordinates"
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCR"
blnResult = VerifyGetVisibleText(objMobiPicker, True)

'Step : Execute GetVisibleText method without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute method without  co-ordinates"
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCR."
blnResult = VerifyGetVisibleText(objMobiPicker, False)


'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object refresh"
Environment("ExpectedResult") = "RefreshObject should re-identify  the object in the application"
blnResult = VerifyRefreshObject(objMobiPicker)


'Step  : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify object  type and class"
Environment("ExpectedResult") = "ToString should return the object type and class."
blnResult = VerifyTOString(objMobiPicker)

'Step   : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify waitproperty when object is visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value"
blnResult = VerifyWaitProperty(objMobiPicker, "visible", True , 5000, True)

'Navigate to other screen
MobiDevice("UICatalog").MobiButton("btnBack").Click
Wait 3

'Step  : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify waitproperty when object is not visible"
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and return False"
blnResult = VerifyWaitProperty(objMobiPicker, "visible",True, 5000, False)

'Navigate back to object screen
NavigateToObjectScreenUICatalog  objMobiPicker ,  , "pickers" , "Pickers"

'Step  : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify property values after update"
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnResult = VerifySetTOProperty(objMobiPicker, arrTOProps)

'Step  : Execute GetItem without parameter
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem without parameter"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker ,  ,  , , "withoutparameter")

'Step  : Execute GetItem with first index only  
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with first index only"
Environment("ExpectedResult") = "Correct index value should be returned"
blnResult = VerifyGetItem(objMobiPicker  , 0  ,  , "John Appleseed" , "withindexonly")


'Step  : Execute GetItem with last  index only  and wheel 0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with last index and wheel 0"
Environment("ExpectedResult") = "Correct index value should be returned"
blnResult = VerifyGetItem(objMobiPicker  , 6  , 0 , "Alain Briere" , "withindexonly") 

'Step  : Execute GetItem with negative index and wheel 0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with negative index value and wheel 0"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker , -20  , 0   ,  "John Appleseed" , "withindexonly")


'Step  : Execute GetItem with out of range index and wheel zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with out of range index and wheel 0"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker , 20  , 0 ,   "John Appleseed" , "withindexonly")

'Step  : Execute GetItem with first index only  and wheel  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with first index  and wheel 1"
Environment("ExpectedResult") = "Correct index value should be returned"
blnResult = VerifyGetItem(objMobiPicker  , 0  , 1 , "0" , "withbothparameters")


'Step  : Execute GetItem with last  index only  and wheel  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with last index and wheel 1"
Environment("ExpectedResult") = "Correct index value should be returned"
blnResult = VerifyGetItem(objMobiPicker  , 6  , 1 , "6" , "withbothparameters") 

'Step  : Execute GetItem with negative index and wheel 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with negative index value and wheel 1"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker , -20  , 1   ,  "0" , "withindexonly")


'Step  : Execute GetItem with out of range index and wheel  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with out of range index and wheel 1"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker , 20  , 1,   "1" , "withindexonly")


'Step  : Execute GetItem with  index value as string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with index value passed as String"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker , "2"  ,  , , "withindexonly")

'Step  : Execute GetItem with  only wheelcount 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with only wheelcount"
Environment("ExpectedResult") = "Error should be thrown"
blnResult = VerifyGetItem(objMobiPicker ,   , 1 , , "withonlyoneparameter")

'Step  : Execute GetItem with  both parameters for wheel 0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with both valid Index and Wheelcount 0"
Environment("ExpectedResult") = "Correct value should get returned"
blnResult = VerifyGetItem(objMobiPicker  ,  2 , 0 , "Serena Auroux" , "withbothparameters")

'Step  : Execute GetItem with  both parameters for wheel  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with both valid Index and Wheelcount  1"
Environment("ExpectedResult") = "Correct value should get returned"
blnResult = VerifyGetItem(objMobiPicker  ,  4 , 1 , "4" , "withbothparameters")

'Step  : Execute GetItem with  negative wheelcount
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify GetItem with negative wheelcount"
Environment("ExpectedResult") = "Error message should be thrown"
blnResult = VerifyGetItem(objMobiPicker  ,  1 , -1 , "" , "withbothparameters")

'Step  : Execute ScrolledText  with  valid wheel as 0 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify text returned after scrolling for wheel 0"
Environment("ExpectedResult") = "Entire Picker text should be returned"
strText = "John Appleseed,Chris Armstrong,Serena Auroux,Susan Bean,Luis Becerra,Kate Bell,Alain Briere"
blnResult = VerifyGetScrollText(objMobiPicker   , strText , True  , 0 , "withvalidwheel")

'Step  : Execute ScrolledText  with  valid wheel as 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify text returned after scrolling for wheel 1"
Environment("ExpectedResult") = "Entire Picker text should be returned"
strText = "0,1,2,3,4,5,6"
blnResult = VerifyGetScrollText(objMobiPicker   , strText , True  , 1 , "withvalidwheel")

'Step  : Execute ScrolledText  without wheelcount
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify text returned after scrolling"
Environment("ExpectedResult") = "Entire Picker text should be returned"
blnResult = VerifyGetScrollText(objMobiPicker  , "" , True , 1 , "withoutwheelcount")

'Step  : Execute ScrolledText  with negative wheel 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method with negative wheelcount"
Environment("ExpectedResult") = "Error message should be thrown"
blnResult = VerifyGetScrollText(objMobiPicker  , "" , True , -20 ,  "withnegativewheel")

'Step  : Execute ScrolledText  with out of index wheel
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method with out of index wheel"
Environment("ExpectedResult") = "Error message should be thrown"
blnResult = VerifyGetScrollText(objMobiPicker  ,  "" , True , 20 , "outofindexwheel")

'Step  : Execute ScrolledText  with  string wheelcount
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method with String wheelcount"
Environment("ExpectedResult") = "Error message should be thrown"
blnResult = VerifyGetScrollText(objMobiPicker  , "" , True , "2" , "stringvaluepassed")

'Step  : Execute RowCount  
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify number of rows returned"
Environment("ExpectedResult") = "Correct row count should be returned"
blnResult = VerifyRowCount(objMobiPicker , 7 , 1)

'Step  : Execute  Scroll  without parameter
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify message on execution with no parameters"
Environment("ExpectedResult") = "Correct error message should be displayed"
blnResult = VerifyScroll(objMobiPicker , "withoutparameter" , Null)

'Step  : Execute  Scroll  with eBottom
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with only direction eBottom"
Environment("ExpectedResult") = "Correct object after scroll should be displayed"

Set objBottomWheelZero = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerBottom_WheelZero")
blnResult = VerifyScroll(objMobiPicker , "bottom" , objBottomWheelZero)

'Step  : Execute  Scroll  with eTop
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with only direction eTOP"
Environment("ExpectedResult") = "Correct object after scroll should be displayed"

Set objTopWheelZero = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerTop_WheelZero")
blnResult = VerifyScroll(objMobiPicker , "top" , objTopWheelZero)

'Step  : Execute  Scroll  with eBottom and wheel Zero 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with direction eBottom and Wheel 0 "
Environment("ExpectedResult") = "Correct object after scroll should be displayed"

Set objBottomWheelZero = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerBottom_WheelZero")
blnResult = VerifyScroll(objMobiPicker , "bottomwithwheelZero" , objBottomWheelZero)

'Step  : Execute  Scroll  with eTop and wheel zero
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with only direction eTOP and wheel  0"
Environment("ExpectedResult") = "Correct object after scroll should be displayed"

Set objTopWheelZero = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerTop_WheelZero")
blnResult = VerifyScroll(objMobiPicker , "topwithwheelZero" , objTopWheelZero)


'Step  : Execute  Scroll  with eBottom and wheel  One 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with direction eBottom and Wheel 1"
Environment("ExpectedResult") = "Correct object after scroll should be displayed"

Set objBottomWheelOne = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerBottom_WheelOne")
blnResult = VerifyScroll(objMobiPicker , "bottomwithwheelZero" , objBottomWheelOne)

'Step  : Execute  Scroll  with eTOP and wheel  One 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with direction eTOP and Wheel 1"
Environment("ExpectedResult") = "Correct object after scroll should be displayed"

Set objTopWheelOne  = MobiDevice("UICatalog").MobiPicker("UIPicker").MobiList("List").MobiElement("PickerTop_WheelOne")
blnResult = VerifyScroll(objMobiPicker , "bottomwithwheelZero" , objBottomWheelOne)

'Step  : Execute  Scroll  Bottom with negative wheel
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with direction bottom and negative wheel"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifyScroll(objMobiPicker , "bottomwithnegativewheel" , Null)

'Step  : Execute  Scroll  Top  with negative wheel
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify scroll with direction top and negative  wheel"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifyScroll(objMobiPicker , "topwithnegativewheel" , Null)

'Step : Execute Swipe without parameters
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe without parameters"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifySwipe(objMobiPicker ,  ,   ,  ,  , Null)

'Step : Execute Swipe with direction eDOWN
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with only direction eDOWN"
Environment("ExpectedResult") = "Correct object should be displayed"
blnResult = VerifySwipe(objMobiPicker , eDOWN , eSLOW  ,  ,  , objBottomWheelZero)

'Step : Execute Swipe with direction eUP
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with only direction eUP"
Environment("ExpectedResult") = "Correct object should be displayed"
blnResult = VerifySwipe(objMobiPicker , eUP , eFAST  ,  ,  , objTopWheelZero)

'Step : Execute Swipe with direction eDOWN  and Wheel 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN and wheelcount as 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber") = 1
blnResult = VerifySwipe(objMobiPicker , eDOWN , eMEDIUM  ,  ,  , objBottomWheelOne)

'Step : Execute Swipe with direction eUP  and Wheel 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with  direction eUP and wheelcount as 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber") = 1
blnResult = VerifySwipe(objMobiPicker , eUP ,  eSLOW ,  ,  , objTopWheelOne)

'Step : Execute Swipe with direction eUP   and  negative  start percentage
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP and negative start percentage"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifySwipe(objMobiPicker , eUP ,   , -20  ,  , Null)

'Step : Execute Swipe with direction eDOWN  and  negative  End  percentage
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with  direction eDOWN and negative end percentage"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifySwipe(objMobiPicker , eDOWN ,   , 20  , -70 , Null)

'Step : Execute Swipe with direction eUP  and  negative Wheelcount
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with  direction eUP and  negative wheelcount"
Environment("ExpectedResult") = "Error message should be displayed"
Environment("WheelNumber") =  -1
blnResult = VerifySwipe(objMobiPicker , eDOWN ,   , 20  , 70 , Null)


'Step : Execute Swipe with direction eUP  and WheelCount out of range
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with  direction eUP and wheelcount out of range"
Environment("ExpectedResult") = "Error message should be displayed"
Environment("WheelNumber") = 10
blnResult = VerifySwipe(objMobiPicker , eUP ,   , 20  , 70 , Null)

'Step : Execute Swipe with direction eDOWN  , velocity  eFAST AND  wheelcount  0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eFAST and wheelcount 0"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 0 
blnResult = VerifySwipe(objMobiPicker , eDOWN , eFAST  ,   ,  , objBottomWheelZero)

'Step : Execute Swipe with direction eUP  , velocity  eFAST AND  wheelcount  0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eFAST and wheelcount 0"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 0 
blnResult = VerifySwipe(objMobiPicker , eUP , eFAST  ,   ,  , objTopWheelZero)

'Step : Execute Swipe with direction eDOWN  , velocity  eFAST AND  wheelcount  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eFAST and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eDOWN , eFAST  ,   ,  , objBottomWheelOne)

'Step : Execute Swipe with direction eUP  , velocity  eFAST AND  wheelcount  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eFAST and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eUP , eFAST  ,   ,  , objTopWheelOne)

'Step : Execute Swipe with direction eDOWN  , velocity  eMEDIUM AND  wheelcount  0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eMEDIUM and wheelcount 0"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 0 
blnResult = VerifySwipe(objMobiPicker , eDOWN , eMEDIUM  ,   ,  , objBottomWheelZero)

'Step : Execute Swipe with direction eUP  , velocity  eMEDIUM AND  wheelcount  0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eMEDIUM and wheelcount 0"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 0 
blnResult = VerifySwipe(objMobiPicker , eUP , eMEDIUM  ,   ,  , objTopWheelZero)

'Step : Execute Swipe with direction eDOWN  , velocity  eMEDIUM  AND  wheelcount  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eMEDIUM and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eDOWN , eMEDIUM  ,   ,  , objBottomWheelOne)

'Step : Execute Swipe with direction eUP  , velocity  eSLOW AND  wheelcount  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eSLOW and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eUP , eSLOW  ,   ,  , objTopWheelOne)

'Step : Execute Swipe with direction eDOWN  , velocity  eSLOW AND  wheelcount  0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eSLOW and wheelcount 0"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 0 
blnResult = VerifySwipe(objMobiPicker , eDOWN , eSLOW  ,   ,  , objBottomWheelZero)

'Step : Execute Swipe with direction eUP  , velocity  eSLOW AND  wheelcount  0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eSLOW and wheelcount 0"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 0 
blnResult = VerifySwipe(objMobiPicker , eUP , eSLOW  ,   ,  , objTopWheelZero)

'Step : Execute Swipe with direction eDOWN  , velocity  eSLOW  AND  wheelcount  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eSLOW and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eDOWN , eSLOW  ,   ,  , objBottomWheelOne)

'Step : Execute Swipe with direction eUP  , velocity  eSLOW AND  wheelcount  1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eSLOW and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eUP , eSLOW  ,   ,  , objTopWheelOne)

'Step : Execute Swipe with direction eDOWN  , velocity  eSLOW  ,startperc 20 , endperc 80 and  wheelcount 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eDOWN , velocity eSLOW , start percentage 20 , end percentage 80 and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eDOWN , eSLOW  , 20  ,  80, objBottomWheelOne)

'Step : Execute Swipe with direction eUP  , velocity  eSLOW ,startperc 20 , endperc 80 and  wheelcount 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with direction eUP , velocity eSLOW ,start percentage 20 , end percentage 80 and wheelcount 1"
Environment("ExpectedResult") = "Correct object should be displayed"
Environment("WheelNumber")  = 1
blnResult = VerifySwipe(objMobiPicker , eUP , eSLOW  , 20  ,80  , objTopWheelOne)

'Step : Execute Swipe with invalid direction
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with invalid direction"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifySwipe(objMobiPicker , eABCD , eSLOW  ,   ,  , Null)

'Step : Execute Swipe with invalid  Velocity
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify swipe with invalid direction"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifySwipe(objMobiPicker , eDOWN , eABCD  ,   ,  , Null)

'Step : Execute Select  without  parameter
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select without parameter"
Environment("ExpectedResult") = "Error message should be displayed"
blnResult = VerifySelect(objMobiPicker , "withoutparameter" , "" , Null)

'Step : Execute Select  with string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with string"
Environment("ExpectedResult") = "Item should get selected"
Environment("WheelNumber")  = 0
blnResult = VerifySelect(objMobiPicker  , "selectstring" , "John Appleseed" , objTopWheelZero)

'Step : Execute Select  with string and wheel 0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with string and wheel 0"
Environment("ExpectedResult") = "Item should get selected"
Environment("WheelNumber")  = 0
blnResult = VerifySelect(objMobiPicker  , "selectstring" , "Alain Briere" , objBottomWheelZero)


'Step : Execute Select  with string and wheel 1
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with string and wheel 1"
Environment("ExpectedResult") = "Item should get selected"
Environment("WheelNumber")  = 1
blnResult = VerifySelect(objMobiPicker  , "selectstringwithwheel" , "6" , objBottomWheelOne)

'Step : Execute Select  with string and wheel 0
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify select with string and wheel 1"
Environment("ExpectedResult") = "Item should get selected"
Environment("WheelNumber")  = 1
blnResult = VerifySelect(objMobiPicker  , "selectstringwithwheel" ,  "1"  , objTopWheelOne)
'*********************************************************************************************************************

'End test iteration
EndTestIteration()

















