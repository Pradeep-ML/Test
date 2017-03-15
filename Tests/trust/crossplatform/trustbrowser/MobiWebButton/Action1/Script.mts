'##########################################################################################################
'Objective: Login to the PhoneLookup app and test MobiWebButton
' Test Description: Execute all MobiWebButton methods 
'##########################################################################################################



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
Set objMobiWebButton  = MobiDevice("Web Browser").MobiWebButton("btnTheButton")
Set objGoogle = MobiDevice("Web Browser").MobiWebEdit("edSearch")
Set objGmail = MobiDevice("Web Browser").MobiWebButton("btnSignIn_Gmail")

arrTOProps = Array("visible", "text", "name", "id", "htmlclass" , "enabled")
arrTOPropValues =Array(True, "the button", "button-name", "button-id", "button-class" , True)
arrROProps = Array("htmlclass" , "id")
arrROPropValues = Array("button-class" , "button-id")


'URL of the application to be opened
strURL =  "http://10.10.1.53/qa/ml.html"
strURL1 = "www.google.com"
strURL2 = "www.gmail.com"

'Create an html report template
CreateReportTemplate()
'#######################################################

' Step1: Open Web View
'Expected Result: WebView  should get opened with  desired URL.
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Open WebView with desired URL"
Environment("ExpectedResult") = "URL should get opened up"

If Not  objMobiWebButton.Exist(3) Then
		'Open URL for testing
	OpenURL strURL , objMobiWebButton  , 3
End If


'''*********************************************************************************************************************
' Step2:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .png image"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .png format"
blnResult = VerifyCaptureBitmap(objMobiWebButton , "png")

' Step3:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .bmp file"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiWebButton , "bmp")

' Step 4:  Execute CaptureBitmap to override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify override message for already existing .bmp  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebButton , "override_bmp")

' Step 5:  Execute CaptureBitmap to override .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify override message for already existing .png  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebButton , "override_png")

' Step 6:  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check property when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return True"
blnResult = VerifyCheckProperty(objMobiWebButton, "visible", "True", 5000 , True)

'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiWebButton, "nonrecursive" , 1 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiWebButton, "recursive" ,1)

'Step 9: Execute Exist  when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify existence when object is visible"
Environment("ExpectedResult")  ="Exist should return True"
blnResult = VerifyExist(objMobiWebButton, True , 5)


'Step 10 Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify run time values of Input properties"
Environment("ExpectedResult")= "Correct run time property should be returned"
blnResult = VerifyGetROProperty(objMobiWebButton , arrROProps , arrROPropValues)

'Step 11: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify returned test object property collection"
Environment("ExpectedResult") = "An collection of properties used for object identification should be returned" 
blnResult = VerifyGetTOProperties(objMobiWebButton, arrTOProps)

'Step 12: Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify values used for object identification"
Environment("ExpectedResult") = "Returned  values should be mapped with Input values"
blnResult =  VerifyGetTOProperty(objMobiWebButton, arrTOProps,arrTOPropValues)

'Step 13: Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify object refresh"
Environment("ExpectedResult") = "Object should get refreshed"
blnResult = VerifyRefreshObject(objMobiWebButton)

'Scroll page to send  object out of view
MobiDevice("Web Browser").Swipe eDOWN ,  eFAST , 20 ,80
Wait 3

'Step 14 : Execute ScrollIntoView
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method brings object in view"
Environment("ExpectedResult") =  "Object should be visible at the top of the page"
blnResult = VerifyScrollIntoView(objMobiWebButton)



'Step 15  Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify value set for the specified identification property"
Environment("ExpectedResult") = "Property value should get updated"
blnResult = VerifySetTOProperty(objMobiWebButton, arrTOProps)

'Step 16: Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verify name of the object"
Environment("ExpectedResult") = "String value cointaining the object description should be returned"
blnResult = VerifyTOString(objMobiWebButton)

'Step 17 : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Wait for the property  till timeout to attain value"
Environment("ExpectedResult") ="WaitProperty should return true when object is visible"
blnResult = VerifyWaitProperty(objMobiWebButton, "visible", "True", 5000,True)


'Step 28: EvaluateScript  (For verifying text value of MobiWebButton)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript shouldr eturn text content value."
blnResult =  VerifyEvaluateScript(objMobiWebButton , "this.textContent" , True , "the button")
'#############################################################

'Step 29:  EvaluateScript  (For performing click on  MobiWebButton)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript should perform click."
blnResult =  VerifyEvaluateScript(objMobiWebButton , "this.click()" , False , "the button")
'#############################################################

MobiDevice("Web Browser").Swipe eDOWN ,  eFAST , 20 ,80
Wait 3

'Step 21  IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebButton , "withoutcoords" , "occluded")
'#############################################################

'Step 22 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebButton , "withcentervalues" , "occluded")
'#############################################################

MobiDevice("Web Browser").Swipe eUP ,  eFAST , 20 ,80
Wait 3

'Step 23 IsOccluded  (For IsOccluded when object is in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebButton , "withoutcoords" , "notoccluded")
'#############################################################

'Step 24 IsOccluded  (For IsOccluded when object is in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebButton , "withcentervalues" , "notoccluded")
'#############################################################

Set objMobiWebButton = MobiDevice("Web Browser").MobiWebButton("btnGoogleSearch")
OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2

'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiWebButton without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiWebButton, "withoutcoords")

OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2

'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiWebButton, "withrandomcoords")

OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2

'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiWebButton, "withboundarycoordsTopLeft")
OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2


'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiWebButton, "withboundarycoordsTopRight")

OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2

'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiWebButton, "withboundarycoordsBottomLeft")
OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2

'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiWebButton, "withboundarycoordsBottomRight")

OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2


'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiWebButton, "withxvalue")

OpenURL strURL1 , objMobiWebButton  , 3

MobiDevice("Web Browser").MobiWebEdit("edSearch").Click
Wait 1
MobiDevice("Web Browser").Type "images"
Wait 2

'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebButton with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiWebButton, "withyvalue")


'Open alternative URL
 OpenURL strURL1 , objGoogle ,2 
'OpenURL strURL2 , objGmail ,2 

Set objMobiWebButton  = MobiDevice("Web Browser").MobiWebButton("btnTheButton")
'MobiDevice("Web Browser").MobiElement("eleGoogle").WaitProperty  "visible" , True , 7000

'Step 25 : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Wait for the property  till timeout to attain value"
Environment("ExpectedResult") = "WaitProperty should return false when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebButton, "visible", "True", 5000, False)

' Step 26:  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Check property when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return False"
blnResult = VerifyCheckProperty(objMobiWebButton, "visible", "True", 5000 , False)

'Step 27: Execute Exist  when object  is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify existence when object is not visible"
Environment("ExpectedResult")  ="Exist should return False"
blnResult = VerifyExist(objMobiWebButton, False, 5)

'End test iteration
EndTestIteration()















































