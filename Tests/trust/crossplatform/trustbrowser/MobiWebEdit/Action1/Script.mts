'##########################################################################################################
'Objective:Test MobiWebEdit  methods on WebBrowser
' Test Description: Execute all MobiWebEdit methods 
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
Set objMobiWebEdit  =  MobiDevice("Web Browser").MobiWebEdit("edUserName")
Set objMobiWebView = MobiDevice("Web Browser").MobiWebView("WebView")
Set objOnGooglePage = MobiDevice("Web Browser").MobiElement("eleGoogle")

arrTOProps = Array("visible", "name", "maximumlength", "type", "enabled")
arrTOPropValues =Array(True, "uname", 2147483647, "text", True)
arrROProps = Array("text" , "backgroundcolor" , "color")
arrROPropValues = array("" , "rgb(255, 255, 255)" , "rgb(0, 0, 0)")

'URL of the application to be opened
strURL =  "http://10.10.1.53/qa/ml.html"
strURL1 = "www.google.com"


'Create an html report template
CreateReportTemplate()
'#######################################################

' Step: Open Web View
'##########################################################
'Expected Result: WebView  should get opened with  desired URL.
intStep = intStep+1
Environment("StepName") = "Step" & intStep

	If NOT  objMobiWebEdit.Exist(4) Then
			 'Open URL for testing
		OpenURL strURL , objMobiWebEdit , 3 
	
	End If

' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .png image"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .png format"
blnResult = VerifyCaptureBitmap(objMobiWebEdit , "png")

' Step  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .bmp file"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiWebEdit , "bmp")

' Step   Execute CaptureBitmap to override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify override message for already existing .bmp  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebEdit , "override_bmp")

' Step :  Execute CaptureBitmap to override .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify override message for already existing .png  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebEdit , "override_png")

' Step :  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check property when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return True"
blnResult = VerifyCheckProperty(objMobiWebEdit, "visible", "True", 5000 , True)

 'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiWebEdit, "nonrecursive" , 0 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiWebEdit, "recursive" ,0)

Set objMobiWebEdit  =  MobiDevice("Web Browser").MobiWebEdit("edUserName")
'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiWebEdit without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiWebEdit, "withoutcoords")
OpenURL strURL , objMobiWebEdit , 3 


'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiWebEdit, "withrandomcoords")
OpenURL strURL , objMobiWebEdit , 3 


'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiWebEdit, "withboundarycoordsTopLeft")
OpenURL strURL , objMobiWebEdit , 3 


'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiWebEdit, "withboundarycoordsTopRight")
OpenURL strURL , objMobiWebEdit , 3 


'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiWebEdit, "withboundarycoordsBottomLeft")
OpenURL strURL , objMobiWebEdit , 3 


'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiWebEdit, "withboundarycoordsBottomRight")
OpenURL strURL , objMobiWebEdit , 3 


'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiWebEdit, "withxvalue")
OpenURL strURL , objMobiWebEdit , 3 


'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebEdit with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiWebEdit, "withyvalue")
OpenURL strURL , objMobiWebEdit , 3 
Wait 3
'Step 19: EvaluateScript  (For setting value in MobiWebEdit)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript shouldr eturn text content value."
blnResult =  VerifyEvaluateScript(objMobiWebEdit , "this.value='texttobeverified'" , True , "texttobeverified")
'#############################################################

'Step 20  EvaluateScript  (For performing click on  MobiWebEdit)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript should perform click."
blnResult =  VerifyEvaluateScript(objMobiWebEdit , "this.click()" , False , "")
'#############################################################
OpenURL strURL , objMobiWebEdit , 3 

MobiDevice("Web Browser").Swipe eDOWN ,  eFAST , 20 ,80
Wait 3

'Step 21  IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebEdit , "withoutcoords" , "occluded")
'#############################################################

'Step 22 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebEdit , "withcentervalues" , "occluded")
'#############################################################

objMobiWebEdit.ScrollIntoView
Wait 2

''Step 23 IsOccluded  (For IsOccluded when object is  in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiWebEdit , "withoutcoords" , "notoccluded")
''#############################################################
'
''Step 24 IsOccluded  (For IsOccluded when object is  in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiWebEdit , "withcentervalues" , "notoccluded")
''#############################################################



'Step  : Execute  Clear with No text
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   "Verify  method when there is no text in Edit box"
Environment("ExpectedResult") = "Clear shouldn't throw any error message"
strResult = VerifyClear(objMobiWebEdit , "withnotext")

'Step  : Execute  Clear with long string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   "Verify method  for long string containing special characters"
Environment("ExpectedResult") = "Text written in the object should get cleared"
strResult = VerifyClear(objMobiWebEdit , "withlongtext")

'Step  : Execute  Clear with text 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify method  for small string"
Environment("ExpectedResult") = "Text written in the object should get cleared"
strResult = VerifyClear(objMobiWebEdit , "withtext")

'Step : Execute Exist  when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify existence when object is visible"
Environment("ExpectedResult")  ="Exist should return True"
blnResult = VerifyExist(objMobiWebEdit, True , 5)


'Step : Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify run time values of Input properties"
Environment("ExpectedResult")= "Correct run time property should be returned"
wt  = objMobiWebEdit.WaitProperty("text" , "" , 3000)
blnResult = VerifyGetROProperty(objMobiWebEdit , arrROProps , arrROPropValues)

'Step : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify returned test object property collection"
Environment("ExpectedResult") = "An collection of properties used for object identification should be returned" 
blnResult = VerifyGetTOProperties(objMobiWebEdit, arrTOProps)

'Step 14 : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify values used for object identification"
Environment("ExpectedResult") = "Returned  values should be mapped with Input values"
blnResult =  VerifyGetTOProperty(objMobiWebEdit, arrTOProps,arrTOPropValues)

'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify object refresh"
Environment("ExpectedResult") = "Object should get refreshed"
blnResult = VerifyRefreshObject(objMobiWebEdit)

'Step  : Execute ScrollIntoView
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method brings object in view"
Environment("ExpectedResult") =  "Object should be visible at the top of the page"
blnResult = VerifyScrollIntoView(objMobiWebEdit)

'Step : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify value set for the specified identification property"
Environment("ExpectedResult") = "Property value should get updated"
blnResult = VerifySetTOProperty(objMobiWebEdit, arrTOProps)

'Step  : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verify name of the object"
Environment("ExpectedResult") = "String value cointaining the object description should be returned"
blnResult = VerifyTOString(objMobiWebEdit)

'Step : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult") ="WaitProperty should return true when object is visible"
blnResult = VerifyWaitProperty(objMobiWebEdit, "visible", "True", 5000,True)


'Step   : Execute Set with special characters
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method  to set  string containing special characters" 
Environment("ExpectedResult") = "Value should be set"

strToSet = "Testing..~!@#$%^&*()_+{}|:<>?/.,';\][=-`0123456789"
strResult = VerifySet(objMobiWebEdit , strToSet ,null)

'Step  : Execute Set with alphanumeric string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify method to set alphanumeris string"
Environment("ExpectedResult") = "Value should be set"

strToSet = "testing12345string"
strResult = VerifySet(objMobiWebEdit , strToSet ,null)


'Not Required
''Step  : Execute Set with null string
''##########################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") =  "Verify method with null string"
'Environment("ExpectedResult") = "Error message should be thrown"
'
'strResult = VerifySet(objMobiWebEdit , null ,null)
'
'Step  : Execute Set  when object is not in view
'##########################################################
intStep = intStep+1
HideObjectFromView  objMobiWebEdit , objMobiWebView
strToSet = "Object not in view"
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method when edit control is not in view"
Environment("ExpectedResult") = "Value should be set"
strResult = VerifySet(objMobiWebEdit , strToSet , null)

'Step  : Execute  Clear with text  when object is not in view
'##########################################################
intStep = intStep+1
HideObjectFromView  objMobiWebEdit , objMobiWebView
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify method  for small string when edit control is not in view"
Environment("ExpectedResult") = "Text written in the object should get cleared"
strResult = VerifyClear(objMobiWebEdit , "withtext")

'Open alternative URL
 OpenURL strURL1 , objOnGooglePage ,2 
 
' Step :  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Check property when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return False"
blnResult = VerifyCheckProperty(objMobiWebEdit, "visible", "True", 5000 , False)


'Step : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult") = "WaitProperty should return false when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebEdit, "visible", "True", 5000, False)
'###############################################################


'Step : Execute Exist  when object  is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify existence when object is not visible"
Environment("ExpectedResult")  ="Exist should return False"
blnResult = VerifyExist(objMobiWebEdit, False, 5)

'End test iteration
EndTestIteration()


































