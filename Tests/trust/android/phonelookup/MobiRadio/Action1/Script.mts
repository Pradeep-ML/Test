'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiRadio
' Test Description: Execute all MobiRadio methods on Search button

'##########################################################################################################

'#######################################################
'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName

'Create result template
CreateReportTemplate

'#######################################################
'Initializations
intStep = 0
Environment("intStepNo") = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
arrTOProps  = Array("visible" ,"text")
arrTOPropValues = Array("True" , "In Stock")
arrROPropValues =Array("android.widget.RadioButton" , "InStock")
arrROProps =  Array("nativeclass" , "name")

'#######################################################

' Step1: Navigate to Switch object screen
'Expected Result: Switch object screen should be displayed
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Navigate to Search screen"
Environment("ExpectedResult") = "Switch object screen should be displayed"


'Set object for Switch
Set objMobiRadio = MobiDevice("Phone Lookup").MobiRadio("InStock")
Set objMobiRadio1 = MobiDevice("Phone Lookup").MobiRadio("rdoAll")

'Navigate to MobiRadio screen
NavigateScreenOnPhoneLookup "Search"  , objMobiRadio , ""

'*********************************************************************************************************************
' Step 2:  Execute CaptureBitmap with .png extention
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiRadio with .png extention."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file to the defined location with .png extention."
blnFlag = VerifyCaptureBitmap(objMobiRadio,"png")


' Step 3:  Execute CaptureBitmap with .bmp extention
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiRadio with .bmp extention."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file to the defined location with .bmp extention."
blnFlag = VerifyCaptureBitmap(objMobiRadio,"bmp")


' Step 4:  Execute CaptureBitmap with .bmp extention already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiRadio with .bmp extention already exist."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw error message as the image already exist."
blnFlag = VerifyCaptureBitmap(objMobiRadio,"override_bmp")


' Step 5:  Execute CaptureBitmap with .png extention already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap on MobiRadio with .png extention already exist."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw error message as the image already exist."
blnFlag = VerifyCaptureBitmap(objMobiRadio,"override_png")


' Step 6:  Execute 'CheckProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiRadio when object is visible."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strValue = objMobiRadio.GetROProperty("text")
blnFlag = VerifyCheckProperty(objMobiRadio, "visible", "True" , 5000, True)


' Step 13:  Execute Click  with boundary coordinates at Top-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiRadio1.Set
Wait 2
'MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobiRadio, "withboundarycoordsTopLeft")

' Step 13:  Execute Click  with boundary coordinates at Top-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Right corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiRadio1.Set
Wait 2
'MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobiRadio, "withboundarycoordsTopRight")

' Step 13:  Execute Click  with boundary coordinates at Bottom-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Left corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
objMobiRadio1.Set
Wait 2
'MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobiRadio, "withboundarycoordsBottomLeft")

' Step 13:  Execute Click  with boundary coordinates at Bottom-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Right corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
'MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
objMobiRadio1.Set
Wait 2
blnFlag = VerifyClick(objMobiRadio, "withboundarycoordsBottomRight")



' Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiRadio recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiRadio,"recursive",0)

' Step 7:  Execute  ChildObjects nonrecusrively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiRadio non-recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiRadio,"nonrecursive",0)

'Step 10 : Execute Click  with Random Co-ordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with Random Co-ordinates on MobiRadio."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click with Random Co-ordinates on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly for Random Co-ordinates."
objMobiRadio1.Set 
Wait 2
blnFlag = VerifyClick(objMobiRadio, "withrandomcoords")




'Step 12 : Execute Click without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click without co-ordinates on MobiRadio"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
objMobiRadio1.Set
Wait 2
blnFlag = VerifyClick(objMobiRadio, "withoutcoords")




'Step 14 : Execute Click with x coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with x co-ordinates on MobiRadio"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with x coordinates."
objMobiRadio1.Set
Wait 2
blnFlag = VerifyClick(objMobiRadio, "withxvalue")


'Step 15 : Execute Click with y coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with y co-ordinates on MobiRadio"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with y coordinates."
objMobiRadio1.Set
Wait 2
blnFlag = VerifyClick(objMobiRadio, "withyvalue")


'Step 16 : Execute Click with valid coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with valid co-ordinates on MobiRadio"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click without coords on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly with valid coordinates."
objMobiRadio1.Set
Wait 2
blnFlag = VerifyClick(objMobiRadio, "withvalidvalue")


'Step 17 : Execute Exist when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiRadio when object is visible"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly when object is visible."
blnFlag = VerifyExist(objMobiRadio, True, 5)


'Step 19: Execute GetROProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiRadio."

Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnFlag = VerifyGetROProperty(objMobiRadio,arrROProps ,arrROPropValues)

'Step 20: Execute GetTOProperties
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiRadio."
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."

blnFlag = VerifyGetTOProperties(objMobiRadio,arrTOProps)

'Step 21: Execute GetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiRadio."

Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."

blnFlag = VerifyGetTOProperty(objMobiRadio, arrTOProps, arrTOPropValues)


'Step 22: Execute GetVisibleText
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText on MobiRadio with  co-ordinates."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnFlag = VerifyGetVisibleText(objMobiRadio,True)


'Step 23: Execute GetVisibleText  without coordinates
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText without coords on MobiRadio without co-ordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText without coords on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
blnFlag = VerifyGetVisibleText(objMobiRadio,False)


'Step 24 : Execute RefreshObject
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject  on MobiRadio."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RefreshObject  on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnFlag = VerifyRefreshObject(objMobiRadio)


'Step 25: Execute Set  when object is not selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Activate on MobiRadio."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Set should set the Switch to correct state: Activated."

blnFlag = VerifySet(objMobiRadio, "" , objMobiRadio2)

'Step 25: Execute Set  when object is already selected
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Set Activate on MobiRadio when object is already selected"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Set should keep object state to Activate."

blnFlag = VerifySet(objMobiRadio, "" , null)


'Step 31: Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetTOProperty on MobiRadio."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute SetTOProperty on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnFlag = VerifySetTOProperty(objMobiRadio, arrTOProps)


'Step 32: Execute 'ToString
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute TOString on MobiRadio."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute TOString on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnFlag = VerifyTOString(objMobiRadio)


'Step 33: Execute 'WaitProperty when object is visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiRadio when object is visible."

Environment("ExpectedResult") = "WaitProperty should wait for the property to attain True"

blnFlag = VerifyWaitProperty(objMobiRadio, "visible", "True", 5000, True)

NavigateScreenOnPhoneLookup  "Controls" ,  MobiDevice("Phone Lookup").MobiDatetimePicker("DatePicker") , "DatePicker"

' Step 7:  Execute 'CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty on MobiRadio when object is not visible."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."

blnFlag = VerifyCheckProperty(objMobiRadio, "visible", "True" , 5000, False)


'Step 18 : Execute Exist when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist on MobiRadio when object is not visible"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiRadio." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly when object is not visible."
blnFlag = VerifyExist(objMobiRadio, False, 5)


'Step 34: Execute 'WaitProperty when object is not visible
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty on MobiRadio when object is not visible."

Environment("ExpectedResult") = "WaitProperty should wait for the property to attain False"

blnFlag = VerifyWaitProperty(objMobiRadio, "visible", "True", 5000, False)
'*********************************************************************************************************************
'End Test  Iteration
EndTestIteration

