'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiSwitch
' Test Description: Execute all MobiSwitch methods on Search button
' Steps:
' Step1 : Navigate to Switch object screen
' Step2 : Execute CaptureBitmap 
' Step3 : Execute CheckProperty
' Step4: Execute ChildObjects
' Step5: Execute Click with Bounadry Co-ordinates
'Setp6: Execute Click  with Random Co-ordinates
'Step7: Execute Click with Zero Co-ordinates
' Step8: Execute  Click without coordinates
' Step9: Execute Exist
' Step10: Execute GetROProperty
' Step11: Execute GetTOProperties
' Step12: Execute GetTOProperty
' Step13:Execute GetVisibleText
'Step14:Execute GetVisibleText without coordinates
' Step15 Execute RefreshObject
' Step16 Execute Set activate
' Step17 Execute Set activate
' Step18 Execute Set Deactivate
' Step19 Execute Set  Deactivate
' Step 20: Execute SetToProperty
' Step21: Execute ToString
' Step22: Execute WaitProperty
'##########################################################################################################
'##########################################################################################################
'Created By : Saurabh Ahuja
'#######################################################
'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
CreateReportTemplate
'#######################################################

'#######################################################
'Initializations
'Initializations
intStep = 0
Environment("intStepNo") = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""

arrProperty = Array("visible" , "accessibilityidentifier","enabled")
arrPropertyValue = Array(True,"controltogglebutton",True)
'#######################################################
'*****************************************************************************************************************
' Step1: Navigate to Switch object screen
'Expected Result: Switch object screen should be displayed
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Navigate to Search screen"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Navigate to Search screen" & VBNewLine
Environment("ExpectedResult") = "Switch object screen should be displayed"

'Set object for Switch
Set objMobiSwitch = MobiDevice("Phone Lookup").MobiSwitch("Vibrateoff")

NavigateScreenOnPhoneLookup "Controls"  , objMobiSwitch , "ToggleButton" 


'*********************************************************************************************************************
' Step 2:  Execute CaptureBitmap with .png extention
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .png extention."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file to the defined location with .png extention."
blnFlag = VerifyCaptureBitmap(objMobiSwitch,"png")


' Step 3:  Execute CaptureBitmap with .bmp extention
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .bmp extention."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file to the defined location with .bmp extention."
blnFlag = VerifyCaptureBitmap(objMobiSwitch,"bmp")


' Step 4:  Execute CaptureBitmap with .bmp extention already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .bmp extention already exist."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw error message as the image already exist."
blnFlag = VerifyCaptureBitmap(objMobiSwitch,"override_bmp")


' Step 5:  Execute CaptureBitmap with .png extention already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiSwitch with .png extention already exist."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw error message as the image already exist."
blnFlag = VerifyCaptureBitmap(objMobiSwitch,"override_png")


' Step 6:  Execute 'CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiSwitch when object is visible."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and return True"

blnFlag = VerifyCheckProperty(objMobiSwitch, "visible", "True" , 5000, True)

'Step 8 : Execute ChildObjects for recursive
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Child objects"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
strResult = VerifyChildObjects(objMobiSwitch , "recursive" , 1)


'Step 9 : Execute ChildObjects for non-recursive
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Child objects"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiCheckBox." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
strResult = VerifyChildObjects(objMobiSwitch , "nonrecursive" , 1)


'Step 9 : Execute Click  with Boundary Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Boundary Co-ordinates on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Boundary Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for Boundary Co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withboundarycoords")


'Step 10 : Execute Click  with Random Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Random Co-ordinates on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Random Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for Random Co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withrandomcoords")


'Step 11 : Execute Click  with Zero Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Zero Co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Zero Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for zero co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withzerovalues")


'Step 12 : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click without co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withoutcoords")


'Step 13 : Execute Click with negative co-ordinaties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with negative co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed  for negative co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withnegativecoords")


'Step 14 : Execute Click with x coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with x co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with x coordinates."
blnFlag = VerifyClick(objMobiSwitch, "withxvalue")


'Step 15 : Execute Click with y coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with y co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with y coordinates."
blnFlag = VerifyClick(objMobiSwitch, "withyvalue")


'Step 16 : Execute Click with valid coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with valid co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with valid coordinates."
blnFlag = VerifyClick(objMobiSwitch, "withvalidvalue")


'Step 17 : Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiSwitch when object is visible"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly when object is visible."
blnFlag = VerifyExist(objMobiSwitch, True, 5)


'Step 19: Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetROProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
arrProperty = Array("visible" , "accessibilityidentifier","enabled")
arrPropertyValue = Array(True,"controltogglebutton",True)
blnFlag = VerifyGetROProperty(objMobiSwitch,arrProperty,arrPropertyValue)

'Step 20: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperties on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
'arrProps = Array("enabled")
blnFlag = VerifyGetTOProperties(objMobiSwitch,arrProperty)

'Step 21: Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
'arrProps = Array("enabled")
blnFlag = VerifyGetTOProperty(objMobiSwitch, arrProperty, arrPropertyValue)


'Step 22: Execute GetVisibleText
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiSwitch with co-ordinates."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnFlag = VerifyGetVisibleText(objMobiSwitch,True)


'Step 23: Execute GetVisibleText without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText without coords on MobiSwitch without co-ordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnFlag = VerifyGetVisibleText(objMobiSwitch,False)


'Step 24 : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject  on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RefreshObject  on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnFlag = VerifyRefreshObject(objMobiSwitch)


'Step 25: Execute Set Activate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Activate on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

blnFlag = VerifySet(objMobiSwitch, eACTIVATE,null)

'Step 26: Execute Set Activate when already activated
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Activate when already active on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Object state should remain Active"

blnFlag = VerifySet(objMobiSwitch, eACTIVATE,null)


'Step 27 : Execute Set deActivate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set deActivate when already active on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: deActivated."

blnFlag = VerifySet(objMobiSwitch, eDEACTIVATE,null)


'Step 28 : Execute Set deActivate when already deActivated
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set deActivate when already deactive on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Object state should remain deactivated"

blnFlag = VerifySet(objMobiSwitch, eDEACTIVATE, null)



'Step 31: Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetTOProperty on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute SetTOProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnFlag = VerifySetTOProperty(objMobiSwitch, arrProperty)



'Step 32: Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute TOString on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute TOString on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnFlag = VerifyTOString(objMobiSwitch)


'Step 33: Execute 'WaitProperty when object is  visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiSwitch when object is visible."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and return True"
blnFlag = VerifyWaitProperty(objMobiSwitch, "visible", "True", 5000, True)

'Logout to navigate to login screen
LogOut

' Step 7:  Execute 'CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiSwitch when object is not visible."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value andreturn False."

blnFlag = VerifyCheckProperty(objMobiSwitch, "visible", "True" , 5000, False)

'Step 18 : Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiSwitch when object is not visible"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly when object is not visible."
blnFlag = VerifyExist(objMobiSwitch, False, 5)

'Step 34: Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiSwitch when object is not visible."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a valueand return False."

blnFlag = VerifyWaitProperty(objMobiSwitch, "visible", "True", 5000, False)
'*********************************************************************************************************************
'Initialization For Switch

arrProperty = Array("visible" , "accessibilityidentifier","enabled")
arrPropertyValue = Array(True,"controlswitch",True)

'#######################################################
' Step1: Navigate to Switch object screen

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Navigate to Search screen"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Navigate to Search screen" & VBNewLine
Environment("ExpectedResult") = "Switch object screen should be displayed"

'Set object for Switch
Set objMobiSwitch = MobiDevice("Phone Lookup").MobiSwitch("OFF")

NavigateScreenOnPhoneLookup "Controls"  , objMobiSwitch , "Switch" 


'Step 9 : Execute Click  with Boundary Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Boundary Co-ordinates on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Boundary Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for Boundary Co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withboundarycoords")


'Step 10 : Execute Click  with Random Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Random Co-ordinates on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Random Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for Random Co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withrandomcoords")


'Step 11 : Execute Click  with Zero Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Zero Co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Zero Co-ordinates on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for zero co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withzerovalues")


'Step 12 : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click without co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withoutcoords")


'Step 13 : Execute Click with negative co-ordinaties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with negative co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed  for negative co-ordinates."
blnFlag = VerifyClick(objMobiSwitch, "withnegativecoords")


'Step 14 : Execute Click with x coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with x co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with x coordinates."
blnFlag = VerifyClick(objMobiSwitch, "withxvalue")


'Step 15 : Execute Click with y coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with y co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with y coordinates."
blnFlag = VerifyClick(objMobiSwitch, "withyvalue")


'Step 16 : Execute Click with valid coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with valid co-ordinates on MobiSwitch"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with valid coordinates."
blnFlag = VerifyClick(objMobiSwitch, "withvalidvalue")


'Step 25: Execute Set Activate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Activate on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

blnFlag = VerifySet(objMobiSwitch, eACTIVATE,null)


'Step 26: Execute Set Activate when already activated
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Activate when already active on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Object state should remain Active"

blnFlag = VerifySet(objMobiSwitch, eACTIVATE,null)


'Step 27 : Execute Set deActivate
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set deActivate when already active on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: deActivated."

blnFlag = VerifySet(objMobiSwitch, eDEACTIVATE,null)


'Step 28 : Execute Set deActivate when already deActivated
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set deActivate when already deactive on MobiSwitch."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "Object state should remain deactivated"

blnFlag = VerifySet(objMobiSwitch, eDEACTIVATE, null)




'Step 22: Execute GetVisibleText
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiSwitch with co-ordinates."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText on MobiSwitch." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnFlag = VerifyGetVisibleText(objMobiSwitch,True)


'End test iteration
EndTestIteration

