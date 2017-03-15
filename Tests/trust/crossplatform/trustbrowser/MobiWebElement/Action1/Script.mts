'##########################################################################################################
'Objective:Test MobiWebElement  methods on WebBrowser
' Test Description: Execute all MobiWebElement methods 
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

Set objMobiWebElement= MobiDevice("Web Browser").MobiWebElement("eleTable")


'objMobiWebElement.ScrollIntoView 
'Wait 3

arrTOProps = Array("visible", "text", "htmltag", "enabled")
arrTOPropValues =Array(True, "table with only tr/td", "P", True)
arrROProps = Array("text")
arrROPropValues = array("table with only tr/td" )

'URL of the application to be opened
strURL =  "http://10.10.1.53/qa/ml.html"
strURL1 = "www.gmail.com"


'Create an html report template
CreateReportTemplate()
'#######################################################

' Step: Open Web View
'##########################################################
'Expected Result: WebView  should get opened with  desired URL.
intStep = intStep+1
Environment("StepName") = "Step" & intStep

	If NOT  objMobiWebElement.Exist(4) Then
			 'Open URL for testing
		OpenURL strURL , objMobiWebElement , 3 
	
	End If

objMobiWebElement.ScrollIntoView 
Wait 3
' Step:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .png image"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .png format"
blnResult = VerifyCaptureBitmap(objMobiWebElement , "png")

' Step  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .bmp file"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiWebElement , "bmp")

' Step   Execute CaptureBitmap to override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify override message for already existing .bmp  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebElement , "override_bmp")

' Step :  Execute CaptureBitmap to override .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify override message for already existing .png  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebElement , "override_png")

' Step :  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check property when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return True"
blnResult = VerifyCheckProperty(objMobiWebElement, "visible", "True", 5000 , True)

 
'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiWebElement, "nonrecursive" , 1 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiWebElement, "recursive" ,1)


Set objMobiWebElement= MobiDevice("Web Browser").MobiWebElement("eleCreateAccount")
OpenURL strURL1 , objMobiWebElement  , 3

'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiWebElement without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiWebElement, "withoutcoords")

OpenURL strURL1 , objMobiWebElement  , 3
'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiWebElement, "withrandomcoords")
OpenURL strURL1 , objMobiWebElement  , 3

'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiWebElement, "withboundarycoordsTopLeft")
OpenURL strURL1 , objMobiWebElement  , 3



'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiWebElement, "withboundarycoordsTopRight")

OpenURL strURL1 , objMobiWebElement  , 3



'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiWebElement, "withboundarycoordsBottomLeft")
OpenURL strURL1 , objMobiWebElement  , 3



'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiWebElement, "withboundarycoordsBottomRight")

OpenURL strURL1 , objMobiWebElement  , 3



'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiWebElement, "withxvalue")

OpenURL strURL1 , objMobiWebElement  , 3



'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebElement with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiWebElement, "withyvalue")

Set objMobiWebElement= MobiDevice("Web Browser").MobiWebElement("eleTable")
OpenURL strURL , objMobiWebElement  , 3

'Step : Execute Exist  when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify existence when object is visible"
Environment("ExpectedResult")  ="Exist should return True"
blnResult = VerifyExist(objMobiWebElement, True , 5)


'Step : Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify run time values of Input properties"
Environment("ExpectedResult")= "Correct run time property should be returned"
wt  = objMobiWebElement.WaitProperty("text" , "" , 3000)
blnResult = VerifyGetROProperty(objMobiWebElement , arrROProps , arrROPropValues)

'Step : Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify returned test object property collection"
Environment("ExpectedResult") = "An collection of properties used for object identification should be returned" 
blnResult = VerifyGetTOProperties(objMobiWebElement, arrTOProps)

'Step 14 : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify values used for object identification"
Environment("ExpectedResult") = "Returned  values should be mapped with Input values"
blnResult =  VerifyGetTOProperty(objMobiWebElement, arrTOProps,arrTOPropValues)

'Step  : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify object refresh"
Environment("ExpectedResult") = "Object should get refreshed"
blnResult = VerifyRefreshObject(objMobiWebElement)

'Step  : Execute ScrollIntoView
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method brings object in view"
Environment("ExpectedResult") =  "Object should be visible at the top of the page"
blnResult = VerifyScrollIntoView(objMobiWebElement)

'Step : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify value set for the specified identification property"
Environment("ExpectedResult") = "Property value should get updated"
blnResult = VerifySetTOProperty(objMobiWebElement, arrTOProps)

'Step  : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verify name of the object"
Environment("ExpectedResult") = "String value cointaining the object description should be returned"
blnResult = VerifyTOString(objMobiWebElement)

'Step : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult") ="WaitProperty should return true when object is visible"
blnResult = VerifyWaitProperty(objMobiWebElement, "visible", "True", 5000,True)

'##########################################################

'Step 28: EvaluateScript  (For verifying text value of MobiWebButton)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript shouldr eturn text content value."
blnResult =  VerifyEvaluateScript(objMobiWebElement , "this.textContent" , True , "table with only tr/td")
'#############################################################

'Step 29:  EvaluateScript  (For performing click on  MobiWebButton)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript should perform click."
blnResult =  VerifyEvaluateScript(objMobiWebElement , "this.click()" , False , "the button")
'#############################################################

MobiDevice("Web Browser").Swipe eDOWN ,  eFAST , 20 ,80
Wait 3

'Step 21  IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebElement , "withoutcoords" , "occluded")
'#############################################################

'Step 22 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebElement , "withcentervalues" , "occluded")
'#############################################################

MobiDevice("Web Browser").Swipe eUP ,  eFAST , 20 ,80
Wait 3

'Step 23 IsOccluded  (For IsOccluded when object is in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebElement , "withoutcoords" , "notoccluded")
'#############################################################

'Step 24 IsOccluded  (For IsOccluded when object is in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebElement , "withcentervalues" , "notoccluded")
'#############################################################

'Open alternative URL
 OpenURL strURL1 , MobiDevice("Web Browser").MobiWebElement("eleCreateAccount"),2

'Step : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult") = "WaitProperty should return false when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebElement, "visible", "True", 5000, False)
'###############################################################


'Step : Execute Exist  when object  is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify existence when object is not visible"
Environment("ExpectedResult")  ="Exist should return False"
blnResult = VerifyExist(objMobiWebElement, False, 5)


' Step :  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Check property when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return False"
blnResult = VerifyCheckProperty(objMobiWebElement, "visible", "True", 5000 , False)


'End test iteration
EndTestIteration()






































