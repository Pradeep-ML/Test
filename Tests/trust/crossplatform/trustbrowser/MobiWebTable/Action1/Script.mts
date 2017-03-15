

'##########################################################################################################
'Objective: Test  MobiWebTable  methods on Web Browser.
' Test Description: Execute all MobiWebTable methods.
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
Environment("Component") = "Web Browser"
'#######################################################
'Input parameters

'Set object for MobiWebCheckBox
Set objMobiWebTable = MobiDevice("Web Browser").MobiWebTable("tblTable2")
Set  objGoogle = MobiDevice("Web Browser").MobiElement("eleGoogle")
Set objGmail =  MobiDevice("Web Browser").MobiWebButton("btnSignIn_Gmail")

arrTOProps = Array("visible" , "name" , "id" , "htmlclass" , "enabled")
arrTOPropValues = Array("True" , "table1-name" , "table1-id" , "trtdth" , "True")
arrROProps = Array( "name","id")
arrROPropValues = Array ("table1-name", "table1-id")

'URL of the application to be opened
strURL = "http://10.10.1.53/qa/ml.html"
strURL1 = "www.google.com"
strURL2 = "www.gmail.com"

'Create an html report template
CreateReportTemplate()
'###########################################################

' Step: Open Web View
'##########################################################
'Expected Result: WebView  should get opened with  desired URL.
intStep = intStep+1
Environment("StepName") = "Step" & intStep

	 'Open URL for testing
	OpenURL strURL , objMobiWebTable , 3 
'	MobiDevice("Web Browser").Scroll eBOTTOM
'	MobiDevice("Web Browser").MobiWebLink("A").ScrollIntoView
	wait 2

'to bring object into view
MobiDevice("Web Browser").Swipe eDOWN, eMEDIUM ,20,80
wait  5

MobiDevice("Web Browser").Swipe eDOWN,  eMEDIUM, 20 ,80
wait 5

''to bring object into view
'MobiDevice("Web Browser").Swipe eDOWN, eMEDIUM,20,60
'wait 3
'MobiDevice("Web Browser").Swipe eDOWN,  eMEDIUM,20,60
'wait 3

'MobiDevice("Web Browser").Swipe eDOWN,  eMEDIUM,10,40

' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .png image"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .png format"
blnResult = VerifyCaptureBitmap(objMobiWebTable , "png")

' Step:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .bmp file"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiWebTable , "bmp")

' Step :  Execute CaptureBitmap to override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify override message for already existing .bmp  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebTable , "override_bmp")

' Step :  Execute CaptureBitmap to override .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify override message for already existing .png  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebTable , "override_png")

' Step :  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check property when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return True"
blnResult = VerifyCheckProperty(objMobiWebTable, "visible", "True", 5000 , True)


'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiWebTable, "nonrecursive" , 2 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiWebTable, "recursive" ,50)


'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiWebTable without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiWebTable, "withoutcoords")


'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiWebTable, "withrandomcoords")



'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiWebTable, "withboundarycoordsTopLeft")



'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiWebTable, "withboundarycoordsTopRight")



'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiWebTable, "withboundarycoordsBottomLeft")



'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiWebTable, "withboundarycoordsBottomRight")




'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiWebTable, "withxvalue")



'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebTable with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiWebTable, "withyvalue")

'Step 19: EvaluateScript  (For verifying text value of MobiWebTable)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript shouldr eturn text content value."
blnResult =  VerifyEvaluateScript(objMobiWebTable , "this.textContent" , True , "the button")
'#############################################################


MobiDevice("Web Browser").Swipe eDOWN ,  eFAST , 20 ,80
Wait 3

'Step 21  IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebTable , "withoutcoords" , "occluded")
'#############################################################

'Step 22 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebTable , "withcentervalues" , "occluded")
'#############################################################

objMobiWebTable.ScrollIntoView
Wait 3

'Step 23 IsOccluded  (For IsOccluded when object is in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebTable , "withoutcoords" , "notoccluded")
'#############################################################

'Step 24 IsOccluded  (For IsOccluded when object is in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebTable , "withcentervalues" , "notoccluded")
'#############################################################



'Step : Execute Exist  when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify existence when object is visible"
Environment("ExpectedResult")  ="Exist should return True"
blnResult = VerifyExist(objMobiWebTable, True , 5)


'Step : Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify run time values of Input properties"
Environment("ExpectedResult")= "Correct run time property should be returned"
blnResult = VerifyGetROProperty(objMobiWebTable , arrROProps , arrROPropValues)

'Step : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify returned test object property collection"
Environment("ExpectedResult") = "An collection of properties used for object identification should be returned" 
blnResult = VerifyGetTOProperties(objMobiWebTable, arrTOProps)

'Step  : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify values used for object identification"
Environment("ExpectedResult") = "Returned  values should be mapped with Input values"
blnResult =  VerifyGetTOProperty(objMobiWebTable, arrTOProps,arrTOPropValues)

'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify object refresh"
Environment("ExpectedResult") = "Object should get refreshed"
blnResult = VerifyRefreshObject(objMobiWebTable)

'Step  : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify value set for the specified identification property"
Environment("ExpectedResult")  = "Property value should get updated"
blnResult = VerifySetTOProperty(objMobiWebTable, arrTOProps)

'Step  : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verify name of the object"
Environment("ExpectedResult") = "String value cointaining the object description should be returned"
blnResult = VerifyTOString(objMobiWebTable)


'Step : Execute 'RowCount  
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify returned Row count"
Environment("ExpectedResult") = "RowCount  should return the number of rows in webtable"
blnResult = VerifyRowCount(objMobiWebTable , 4  , "")
'###############################################################

'Step : Execute 'ColumnCount
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify returned Column count "
Environment("ExpectedResult") = "ColumnCount  should return the number of columns in webtable"
blnResult =VerifyColumnCount (objMobiWebTable , 3)
'###############################################################

'Step : Execute VerifyGetRowWithCellText  with valid parameters
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Text returned from method on passing valid text , column index and row index values"
Environment("ExpectedResult")  = "Row index should be returned"
blnResult  = VerifyGetRowWithCellText(objMobiWebTable, "valid", "January", 0, 1)


'Step : Execute VerifyGetRowWithCellText  with valid text only
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Text returned from method on passing only valid text"
Environment("ExpectedResult")  = "Row index should be returned"
blnResult  = VerifyGetRowWithCellText(objMobiWebTable, "validwithonlytext", "January", "" , "")
'###############################################################

'Step : Execute VerifyGetRowWithCellText  with valid text  and column index
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Text returned from method on passing  valid text and column index"
Environment("ExpectedResult")  = "Row index should be returned"
blnResult  = VerifyGetRowWithCellText(objMobiWebTable, "validwithtextandcolumn", "January", 0 , "")
'##########################################################


''Step : Execute VerifyGetRowWithCellText  with column value out of index
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify Text returned from method on passing out of  range column index"
'Environment("ExpectedResult")  = "Error message should be thrown"
'blnResult  = VerifyGetRowWithCellText(objMobiWebTable, "withcolumnoutofindex", "", "" , "")
'##########################################################


''Step : Execute VerifyGetRowWithCellText  with row out of  index
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify Text returned from method on passing  out of range row index"
'Environment("ExpectedResult")  = "Error message should be thrown"
'blnResult  = VerifyGetRowWithCellText(objMobiWebTable, "withrowoutofindex", "", , "")
'##########################################################


'Step : Execute VerifyGetCellData  with  valid  row and coulmn index
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Text returned from method on passing valid column and row index"
Environment("ExpectedResult")  = "Text returned should be mapped with expected value"

strValue = "Month,Savings,Spendings,Sum,$180,$280,January,$100,$200,February,$80,$180"
blnResult  =VerifyGetCellData(objMobiWebTable , strValue , "validvalues")
'##########################################################

''Step : Execute VerifyGetCellData  with  invalid row index
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify Text returned from method on passing invalid row index"
'Environment("ExpectedResult") = "Error message should be thrown"
'
'strValue = 20
'blnResult  =VerifyGetCellData(objMobiWebTable , strValue , "invalidrowindex")
'##########################################################

''Step : Execute VerifyGetCellData  with  invalid column index
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify Text returned from method on passing invalid column index"
'Environment("ExpectedResult")  = "Error message should be thrown"
'
'strValue = 20
'blnResult  =VerifyGetCellData(objMobiWebTable , strValue , "invalidcolumnindex")
'##########################################################

''Step : Execute VerifyGetCellData  without parameters
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify Text returned from method without passing any parameter"
'Environment("ExpectedResult")  = "Error message should be thrown"
'
'blnResult  = VerifyGetCellData(objMobiWebTable , "" , "withoutparameter")
'##########################################################

'Step  : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult")  ="WaitProperty should return true when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebTable, "visible", "True", 5000,True)

'Open alternative URL
 OpenURL strURL1 , objGoogle ,2 
 OpenURL strURL2  , objGmail , 3

'Step : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult")  = "WaitProperty should return false when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebTable, "visible", "True", 5000, False)
'###############################################################

' Step:  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Check property when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return False"
blnResult = VerifyCheckProperty(objMobiWebTable, "visible", "True", 5000 , False)

'Step : Execute Exist  when object  is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify existence when object is not visible"
Environment("ExpectedResult")  ="Exist should return False"
blnResult = VerifyExist(objMobiWebTable, False, 5)

'###########################################################
'End test iteration
EndTestIteration()




















































 




















