
'##########################################################################################################
'Objective: Test MobiWebRadioButton methods on Web Browser
' Test Description: Execute all MobiWebRadioButton methods 
'#################################################################################################

			
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

'Input parameters
Set objMobiWebRadio = MobiDevice("Web Browser").MobiWebRadioButton("rdoDinosaurs")
Set objMobiWebRadio2 = MobiDevice("Web Browser").MobiWebRadioButton("rdoAndroids")
Set objOnGooglePage = MobiDevice("Web Browser").MobiElement("eleGoogle")
Set objGmail =  MobiDevice("Web Browser").MobiWebButton("btnSignIn_Gmail")

arrTOProps = Array("visible" , "value"  , "name" , "id" ,  "htmlclass" , "enabled")
arrTOPropValues = Array("True" ,"dinos"  ,  "radioname" , "dinos-id" , "dinos-class", True)

arrROProps = Array( "value" , "id")
arrROPropValues = Array( "dinos" , "dinos-id")

'URL of the application to be opened
strURL =  "http://10.10.1.53/qa/ml.html"
strURL1 = "www.google.com"
strURL2 = "www.gmail.com"

'Create an html report template
CreateReportTemplate()
'###########################################################

' Step1: Open Web View
'##########################################################
'Expected Result: WebView  should get opened with  desired URL.
intStep = intStep+1
Environment("StepName") = "Step" & intStep

If  Not objMobiWebRadio.Exist(2)  Then
		 'Open URL for testing
	OpenURL strURL , objMobiWebRadio, 3
End If

' Step2:  Execute CaptureBitmap with .png format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .png image"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .png format"
blnResult = VerifyCaptureBitmap(objMobiWebRadio , "png")

' Step3:  Execute CaptureBitmap with .bmp format
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify method for capturing .bmp file"
Environment("ExpectedResult") = "CaptureBitMap should capture screeshot in .bmp format"
blnResult = VerifyCaptureBitmap(objMobiWebRadio , "bmp")

' Step 4:  Execute CaptureBitmap to override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify override message for already existing .bmp  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebRadio , "override_bmp")

' Step 5:  Execute CaptureBitmap to override .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify override message for already existing .png  file"
Environment("ExpectedResult") = "CaptureBitMap should display override message"
blnResult = VerifyCaptureBitmap(objMobiWebRadio , "override_png")

' Step 6:  Execute CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Check property when object is visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return True"
blnResult = VerifyCheckProperty(objMobiWebRadio, "visible", "True", 5000 , True)



'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiWebRadio, "nonrecursive" , 0 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiWebRadio, "recursive" ,0)


'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiWebRadioButton without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobiWebRadio, "withoutcoords")



'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobiWebRadio, "withrandomcoords")



'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobiWebRadio, "withboundarycoordsTopLeft")



'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobiWebRadio, "withboundarycoordsTopRight")



'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobiWebRadio, "withboundarycoordsBottomLeft")



'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobiWebRadio, "withboundarycoordsBottomRight")



'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiWebRadio, "withxvalue")


'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiWebRadioButton with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobiWebRadio, "withyvalue")

'Step 19: EvaluateScript  (For verifying text value of MobiWebRadioButton)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript shouldr eturn text content value."
blnResult =  VerifyEvaluateScript(objMobiWebRadio , "document.getElementById('dinos-id').value" , True , "dinos")
'#############################################################

'Step 20  EvaluateScript  (For performing click on  MobiWebRadioButton)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify EvaluateScript" 
Environment("ExpectedResult") = "EvaluateScript should perform click."
blnResult =  VerifyEvaluateScript(objMobiWebRadio , "this.click()" , False , "")
'#############################################################

MobiDevice("Web Browser").Swipe eDOWN ,  eFAST , 20 ,80
Wait 3
'Step 21  IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebRadio , "withoutcoords" , "occluded")
'#############################################################

'Step 22 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebRadio , "withcentervalues" , "occluded")
'#############################################################

MobiDevice("Web Browser").Swipe eUP ,  eFAST , 20 ,80
Wait 3

'Step 23 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebRadio , "withoutcoords" , "notoccluded")
'#############################################################

'Step 24 IsOccluded  (For IsOccluded when object is not in view)
'###############################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify IsOccluded" 
Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
blnResult =  VerifyIsOccluded(objMobiWebRadio , "withcentervalues" , "notoccluded")
'#############################################################




'Step 10: Execute Exist  when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify existence when object is visible"
Environment("ExpectedResult")  ="Exist should return True"
blnResult = VerifyExist(objMobiWebRadio, True , 5)

'Step 12: Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify run time values of Input properties"
Environment("ExpectedResult")= "Correct run time values for the Input properties should be returned"
blnResult = VerifyGetROProperty(objMobiWebRadio , arrROProps , arrROPropValues)

'Step 13: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify returned test object property collection"
Environment("ExpectedResult") = "An collection of properties used for object identification should be returned" 
blnResult = VerifyGetTOProperties(objMobiWebRadio, arrTOProps)

'Step 14 : Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify values used for object identification"
Environment("ExpectedResult") = "Returned  values should be mapped with Input values"
blnResult =  VerifyGetTOProperty(objMobiWebRadio, arrTOProps,arrTOPropValues)

'Step 15 : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  " Verify object refresh"
Environment("ExpectedResult") = "Object should get refreshed"
blnResult = VerifyRefreshObject(objMobiWebRadio)

'Step 16 : Execute ScrollIntoView
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify method brings object in view"
Environment("ExpectedResult") =  "Object should be visible at the top of the page"
blnResult = VerifyScrollIntoView(objMobiWebRadio)

'Step 17 : Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Verify value set for the specified identification property"
Environment("ExpectedResult") = "Property value should get updated"
blnResult = VerifySetTOProperty(objMobiWebRadio, arrTOProps)

'Step 18 : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") ="Verify name of the object"
Environment("ExpectedResult") = "String value cointaining the object description should be returned"
blnResult = VerifyTOString(objMobiWebRadio)

'Step 19 : Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify Wait  till timeout to attain value for the specified property"
Environment("ExpectedResult") ="WaitProperty should return true when object is  visible"
blnResult = VerifyWaitProperty(objMobiWebRadio, "visible", "True", 5000,True)


'Step : Execute 'Set  when object is not selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Set when object is not selected"
Environment("ExpectedResult") = "Set should select the object"
blnResult = VerifySet(objMobiWebRadio ,  null  , objMobiWebRadio2)
'###############################################################

'Step : Execute 'Set  when object is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify Set when object is already selected"
Environment("ExpectedResult") = "Object state should not change"
blnResult = VerifySet(objMobiWebRadio , null , null)
'###############################################################

OpenURL  strURL1 , objOnGooglePage , 2
OpenURL  strURL2 , objGmail , 2
MobiDevice("Web Browser").Sync

' Step 7:  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Check property when object is not visible"
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a and return False"
blnResult = VerifyCheckProperty(objMobiWebRadio, "visible", "True", 5000 , False)

'Step 11: Execute Exist  when object  is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Verify existence when object is not visible"
Environment("ExpectedResult")  ="Exist should return False"
blnResult = VerifyExist(objMobiWebRadio, False, 5)
MobiDevice("Web Browser").MobiElement("eleGoogle").WaitProperty  "visible" , True , 7000


'Step : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Verify Wait  till timeout to attain value for the specified property"
Environment("ExpectedResult") = "WaitProperty should return false when object is not visible"
blnResult = VerifyWaitProperty(objMobiWebRadio, "visible", "True", 5000, False)


'End test iteration
EndTestIteration()




