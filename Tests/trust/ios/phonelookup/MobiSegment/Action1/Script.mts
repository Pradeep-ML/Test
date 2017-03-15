'##########################################################################################################
' Objective: Launch  PhoneLookup app and test MobiSegment  Object
' Test Description: Execute all methods for MobiSegment  

'##########################################################################################################
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

' Step1: Navigate to Search screen
'Expected Result: Search screen should be displayed
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Verify Login Screen" & VBNewLine
Environment("ExpectedResult") = "Login Screen should be displayed"

'Set object for Button
Set objMobiSegment = MobiDevice("PhoneLookup").MobiSegment("Segment")

'Call function to createreporttemplare
CreateReportTemplate()

'Call navigate to screen function 
strResult  =  LoginAndNavigateToControlsPage ("UISegmentedControl" , objMobiSegment)


' Step2:  Execute CaptureBitmap with .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the png file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiSegment , "png")

' Step3:  Execute CaptureBitmap with .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute capture bitmap with .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the bmp file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiSegment , "bmp")

' Step4:  Execute CaptureBitmap  with override .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "excute capture bitmap with override .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Error message for override of bmp image should appear."
strResult = VerifyCaptureBitmap(objMobiSegment , "override_bmp")

' Step5:  Execute CaptureBitmap with override .png file 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute capture bitmap with override .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiSegment." & VBNewLine
Environment("ExpectedResult") = " Error message for override of png image should appear."
strResult = VerifyCaptureBitmap(objMobiSegment , "override_png")

' Step 6:  Execute 'CheckProperty when object is  visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strChecked  = objMobiSegment.GetROProperty("itemscount")

strResult = VerifyCheckProperty(objMobiSegment, "itemscount", strChecked, 5000, True)

'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiSegment, "nonrecursive" , 3)

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiSegment, "recursive" ,9)


'Step 16 Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = VerifyExist(objMobiSegment, True, 5)

'Step 18 : Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetRoProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."

arrProps = Array("visible", "itemscount","enabled","name")
arrvalue = Array(True ,"3","True","Segment")
strResult = VerifyGetROProperty(objMobiSegment, arrProps, arrvalue)



'Step 19: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetToProperties"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProperties = Array("Visible","itemscount","enabled","allitems")
strResult = VerifyGetTOProperties(objMobiSegment,arrProperties)


'Step 20: Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("visible", "itemscount","enabled","allitems")
arrvalue = Array(True ,"3","True","Check;Search;Tools")
strResult = VerifyGetTOProperty(objMobiSegment, arrProps, arrvalue)


'Step 21  Execute RefreshObject
MobiDevice("PhoneLookup").MobiSegment("Segment").Click

'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Refresh object"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiSegment)

'Step 22 : Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
strResult = VerifyTOString(objMobiSegment)


'Step 23 : Execute 'WaitProperty when object is visible 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
'strChecked  = objMobiSegment.GetROProperty("itemscount")

strResult = VerifyWaitProperty(objMobiSegment, "visible", True, 5000, True)


'Step 25: Execute 'Select  with index value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with index value"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Select should select the passed in text correctly."
Set objMobiSegment = MobiDevice("PhoneLookup").MobiSegment("Segment")
strResult = VerifySelect(objMobiSegment ,"withindex", 0 , objAfterSelection)

'Step 27: Execute 'Select with out of index
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with out of index"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Select should throw error"
Set objMobiSegment = MobiDevice("PhoneLookup").MobiSegment("Segment")
strResult = VerifySelect(objMobiSegment ,"withoutofindex", 5, objAfterSelection)


'Step 28: Execute 'Select with negative index
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with negatvie index"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Select should throw an error"

strResult = VerifySelect(objMobiSegment ,"withnegativeindex",-1 , objAfterSelection)


'Step 29: Execute 'Select index as a  string
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with index as a string"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Error message should be thrown."

strResult = VerifySelect(objMobiSegment ,"withindexasstring", "1" , objAfterSelection)


'Step 30: Execute 'Select hash value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with index as a string"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Set on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Set should set the passed in text correctly."

strResult = VerifySelect(objMobiSegment ,"withhashvalue", "#1" , objAfterSelection)

'Step 31: Execute Getitmes
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute getitmes"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetItems on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "GetItems should return the items correctly."
arrItems = Array ("Check","Search","Tools")
strResult =  VerifyGetItems(objMobiSegment , arrItems)

'Step 32: Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
arrProps = Array("visible", "itemscount","enabled","allitems")
strResult = VerifySetTOProperty(objMobiSegment, arrProps)

''Step 23 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSegment , "withoutcoords" , "notoccluded")
''#############################################################
'
''Step 24 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSegment , "withcentervalues" , "notoccluded")
''#############################################################
'
''Hide MobiSegment object
'MobiDevice("PhoneLookup").MobiEdit("edSearch").Click
'Wait 3
'
''Step 21  IsOccluded  (For IsOccluded when object is not in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view without passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSegment , "withoutcoords" , "occluded")
''#############################################################
'
''Step 22 IsOccluded  (For IsOccluded when object is not in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object is not is view by  passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobiSegment , "withcentervalues" , "occluded")
''#############################################################

'Logout
LogOut

' Step 7  Execute 'CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strResult = VerifyCheckProperty(objMobiSegment, "visible", True, 5000, False)

'Step 17: Execute Exist when object is not visble
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = VerifyExist(objMobiSegment, False, 5)



'Step 24 : Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiSegment." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiSegment, "visible", True , 5000, False)

Login "mobilelabs" , "demo" 

selectMenuOption "Search"

Set objMobiSegment  = MobiDevice("PhoneLookup").MobiSegment("Search_Segment")

'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobiSegment without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
objMobiSegment.Select 0 
blnStepRC = VerifyClick(objMobiSegment, "withoutcoords")


'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
objMobiSegment.Select 1
blnStepRC = VerifyClick(objMobiSegment, "withrandomcoords")


'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
objMobiSegment.Select 2
blnStepRC = VerifyClick(objMobiSegment, "withboundarycoordsTopLeft")


'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
objMobiSegment.Select 0
blnStepRC = VerifyClick(objMobiSegment, "withboundarycoordsTopRight")

'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
objMobiSegment.Select 2
blnStepRC = VerifyClick(objMobiSegment, "withboundarycoordsBottomLeft")

'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
objMobiSegment.Select 0
blnStepRC = VerifyClick(objMobiSegment, "withboundarycoordsBottomRight")


'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobiSegment, "withxvalue")

'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiSegment with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
objMobiSegment.Select 0
blnStepRC = VerifyClick(objMobiSegment, "withyvalue")

'******************************************************************************************************************************************************************

'Call function to end test iteration
EndTestIteration()





