'##########################################################################################################
'Objective:  Test  MobiWebLink methods on Web Browser
' Test Description: Execute all MobiWebLink methods 
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
Set objMobiWebLink  = MobiDevice("Web Browser").MobiWebLink("lnkMobileLabs")

Set objGoogle = MobiDevice("Web Browser").MobiElement("eleGoogle")
Set objGmail =  MobiDevice("Web Browser").MobiWebButton("btnSignIn_Gmail")

arrTOProps = Array("visible" , "text"   , "href", "enabled")
arrTOPropValues = Array("True" ,"Mobile Labs Inc"  ,  "http://www.mobilelabsinc.com/" , "True")
arrROProps = Array("innertext" , "color")
arrROPropValues = Array("Mobile Labs Inc"  , "rgb(0, 0, 238)")


'URL of the application to be opened
strURL =  "http://10.10.1.53/qa/ml.html"
strURL1 = "www.google.com"
strURL2 = "www.gmail.com"

'Create an html report template
CreateReportTemplate()
'###########################################################

' Step1: Open Web View
'Expected Result: WebView  should get opened with  desired URL.
intStep = intStep+1
Environment("StepName") = "Step" & intStep

	 'Open URL for testing
	OpenURL strURL , objMobiWebLink , 3 


' Step2:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .png image"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .png format"
blnResult = VerifyCaptureBitmap(objMobiWebLink , "png")

' Step3:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .bmp file"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiWebLink , "bmp")

' Step 4:  Execute CaptureBitmap to override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify override message for already existing .bmp  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebLink , "override_bmp")

' Step 5:  Execute CaptureBitmap to override .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify override message for already existing .png  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebLink , "override_png")

' Step 6:  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check property when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return True"
blnResult = VerifyCheckProperty(objMobiWebLink, "visible", "True", 5000 , True)


'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiWebLink, "nonrecursive" , 1)

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiWebLink, "recursive" ,1)

 
 
 'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiLink without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiWebLink, "withoutcoords")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3


'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiWebLink, "withrandomcoords")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3


'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiWebLink, "withboundarycoordsTopLeft")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3


'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiWebLink, "withboundarycoordsTopRight")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3


'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiWebLink, "withboundarycoordsBottomLeft")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3


'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiWebLink, "withboundarycoordsBottomRight")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3



'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiWebLink, "withxvalue")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3


'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiLink with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiWebLink, "withyvalue")
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3

'Step 19: EvaluateScript  (For verifying text value of MobiWebLink)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript shouldr eturn text content value."
blnResult =  VerifyEvaluateScript(objMobiWebLink , "this.textContent" , True , "Mobile Labs Inc")
'#############################################################

'Step 20  EvaluateScript  (For performing click on  MobiWebLink)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript should perform click."
blnResult =  VerifyEvaluateScript(objMobiWebLink , "this.click()" , False , "the button")
'#############################################################
'Open URL for testing
 OpenURL strURL , objMobiWebLink ,3
MobiDevice("Web Browser").MobiWebView("WebView").Scroll eBOTTOM
Wait 3

'Step 21  IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebLink , "withoutcoords" , "occluded")
'#############################################################

'Step 22 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebLink , "withcentervalues" , "occluded")
'#############################################################

MobiDevice("Web Browser").Swipe eUP ,  eFAST , 20 ,80
Wait 3

'Step 23 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebLink , "withoutcoords" , "notoccluded")
'#############################################################

'Step 24 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebLink , "withcentervalues" , "notoccluded")
'#############################################################



'Step 10: Execute Exist  when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify existence when object is visible"
Environment("ExpectedResult")  ="Exist should return True"
blnResult = VerifyExist(objMobiWebLink, True , 5)


'Step 12: Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify run time values of Input properties"
Environment("ExpectedResult")= "Correct run time property should be returned"
blnResult = VerifyGetROProperty(objMobiWebLink , arrROProps , arrROPropValues)

'Step 13: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify returned test object property collection"
Environment("ExpectedResult") = "An collection of properties used for object identification should be returned" 
blnResult = VerifyGetTOProperties(objMobiWebLink, arrTOProps)

'Step 14 : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify values used for object identification"
Environment("ExpectedResult") = "Returned  values should be mapped with Input values"
blnResult =  VerifyGetTOProperty(objMobiWebLink, arrTOProps,arrTOPropValues)

'Step 15 : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify object refresh"
Environment("ExpectedResult") = "Object should get refreshed"
blnResult = VerifyRefreshObject(objMobiWebLink)

'Step 16 : Execute ScrollIntoView
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method brings object in view"
Environment("ExpectedResult") =  "Object should be visible at the top of the page"
blnResult = VerifyScrollIntoView(objMobiWebLink)

'Step 17 : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify value set for the specified identification property"
Environment("ExpectedResult") = "Property value should get updated"
blnResult = VerifySetTOProperty(objMobiWebLink, arrTOProps)

'Step 18 : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verify name of the object"
Environment("ExpectedResult") = "String value cointaining the object description should be returned"
blnResult = VerifyTOString(objMobiWebLink)

'Step 19 : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult") ="WaitProperty should return true when object is visible"
blnResult = VerifyWaitProperty(objMobiWebLink, "visible", "True", 5000,True)

'Open alternative URL
 OpenURL strURL1 , objGoogle ,2 
 OpenURL  strURL2  ,  objGmail , 3


'Step : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Wait for the property  till timeout to attain value"
Environment("ExpectedResult") = "WaitProperty should return false when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebLink, "visible", True, 5000, False)


' Step 7:  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Check property when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return False"
blnResult = VerifyCheckProperty(objMobiWebLink, "visible", True, 5000 , False)


'Step 11: Execute Exist  when object  is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify existence when object is not visible"
Environment("ExpectedResult")  ="Exist should return False"
blnResult = VerifyExist(objMobiWebLink, False, 5)
'###############################################################

'End test iteration
EndTestIteration()







































