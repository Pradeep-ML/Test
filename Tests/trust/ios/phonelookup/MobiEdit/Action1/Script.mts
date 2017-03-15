
'##########################################################################################################
' Objective: Login to the PhoneLookup app and in the process test MobiEdit  Methods
' Test Description: Execute all methods for MobiEdit on Username and Password editboxes 
'The methods are: CaptureBitmap, CheckProperty, ChildObjects, 
' Click, Exist, GetROProperty, GetTOProperties, GetTOProperty, GetVisibleText, RefreshObject, SetTOProperty, TOString, 
' WaitProperty , Set

'Steps:
' Step1:  Execute CaptureBitmap
' Step2:  Execute CheckProperty 
' Step3:  Execute  ChildObjects
' Step4:  Execute Clear 
' Step5:  Execute Click  without coordinates
' Step6:  Execute Click  with random coordinates
' Step7:  Execute Click  with boundary coordinates
' Step8:  Execute Click  with zero coordinates
' Step9:  Execute DblClick  without coordinates
' Step10:  Execute DblClick  with random coordinates
' Step11:  Execute DblClick  with boundary coordinates
' Step12:  Execute DblClick  with zero coordinates
' Step13:  Execute Exist 
' Step14:  Execute GetROProperty 
' Step15:  Execute GetTOProperties
' Step16:  Execute GetTOProperty 
' Step17:  Execute GetVisibleText  without coordinates
 'Step18:  Execute GetVisibleText  with coordinates
' Step19:  Execute LongClick  without coordinates
' Step20:  Execute LongClick  with random coordinates
' Step21:  Execute LongClick  with boundary coordinates
' Step22:  Execute LongClick  with zero coordinates
' Step23:  Execute RefreshObject
' Step24:  Execute Set 
' Step25:  Execute SetTOProperty
' Step26:  Execute TOString 
' Step27:  Execute WaitProperty 

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
intStep = 0
Environment("intStepNo") = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
Environment("StepName") = ""
'#######################################################

'#######################################################
'Initial Setup

'Logout if a session is already in progress
Logout

'#######################################################

'Set object for Username edit
Set objMobEdit = MobiDevice("PhoneLookup").MobiEdit("Username")
Set objMobiEdit2 = MobiDevice("PhoneLookup").MobiEdit("Password")

'Activate the MobiEdit
objMobEdit.Click
Window("regexpwndtitle:=deviceViewer").Type micReturn

' Step 1:  Execute CaptureBitmap with .png extension
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "CaptureBitmap : Execute CaptureBitmap with .png extension on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file with .png extension to the defined location."
blnFlag = VerifyCaptureBitmap(objMobEdit,"png")

' Step 2:  Execute CaptureBitmap with .bmp extenstion
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "CaptureBitmap : Execute CaptureBitmap with .bmp extension on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file with .bmp extension to the defined location."
blnFlag = VerifyCaptureBitmap(objMobEdit,"bmp")

' Step 3:  Execute CaptureBitmap with .bmp extenstion already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "CaptureBitmap : Execute CaptureBitmap with .bmp extension already exist on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed"
blnFlag = VerifyCaptureBitmap(objMobEdit,"override_bmp")

' Step 4:  Execute CaptureBitmap with .png extenstion already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "CaptureBitmap : Execute CaptureBitmap with .png extension already exist on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed"
blnFlag = VerifyCaptureBitmap(objMobEdit,"override_png")

' Step 5:  Execute CheckProperty when object is visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "CheckProperty : Execute CheckProperty when object is visible"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
MobiDevice("PhoneLookup").MobiEdit("Username").Set "Go@Mobile"
blnFlag = VerifyCheckProperty(objMobEdit, "text", "Go@Mobile", 5000, True)
MobiDevice("PhoneLookup").MobiEdit("Username").Clear

' Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiEdit recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobEdit,"recursive",2)

' Step 7:  Execute  ChildObjects nonrecusrively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "ChildObjects : Execute ChildObjects on MobiEdit non-recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobEdit,"nonrecursive",1)

' Step 8:  Execute Clear with long text
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Clear : Execute Clear on long text on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Clear on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Clear should clear the text within the editbox."
blnFlag = VerifyClear(MobiDevice("PhoneLookup").MobiEdit("Username"),"withlongtext")

' Step 9:  Execute Clear with text
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Clear : Execute Clear on text on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Clear on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Clear should clear the text within the editbox."
blnFlag = VerifyClear(MobiDevice("PhoneLookup").MobiEdit("Username"),"withtext")

' Step 10:  Execute Clear with no text
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Clear : Execute Clear on no text on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Clear on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Clear should clear the text within the editbox."
blnFlag = VerifyClear(MobiDevice("PhoneLookup").MobiEdit("Username"),"withnotext")

'MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"

' Step 11:  Execute Click  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click without coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withoutcoords")

' Step 12:  Execute Click  with random coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with random coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withrandomcoords")

' Step 13:  Execute Click  with boundary coordinates at Top-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withboundarycoordsTopLeft")

' Step 13:  Execute Click  with boundary coordinates at Top-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Right corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withboundarycoordsTopRight")

' Step 13:  Execute Click  with boundary coordinates at Bottom-Left corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Left corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withboundarycoordsBottomLeft")

' Step 13:  Execute Click  with boundary coordinates at Bottom-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Right corner on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withboundarycoordsBottomRight")

' Step 14:  Execute Click  with zero coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with zero coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withzerovalues")

' Step 15:  Execute Click  with only x coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with only x coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withxvalue")

' Step 16:  Execute Click  with only y coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with only y coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withyvalue")

' Step 17:  Execute Click  with any valid coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with any valid coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyClick(objMobEdit, "withvalidvalue")

' Step 18:  Execute Click  with negative coordinates
'#######################################################

'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Click : Execute Click with negative coordinates on MobiEdit."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Click on MobiEdit." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed for negative values"
'blnFlag = VerifyClick(objMobEdit, "withnegativecoords")

' Step 19:  Execute DblClick  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  without coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withoutcoords")
MobiDevice("PhoneLookup").MobiEdit("Username").Click

' Step 20:  Execute DblClick  with random coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  with random coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withrandomcoords")
MobiDevice("PhoneLookup").MobiEdit("Username").Click

' Step 21:  Execute DblClick  with boundary coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  with boundary coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withboundarycoords")
MobiDevice("PhoneLookup").MobiEdit("Username").Click

' Step 22:  Execute DblClick  with zero coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  with zero coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withzercoords")
MobiDevice("PhoneLookup").MobiEdit("Username").Click


' Step 23:  Execute DblClick  with only x coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  with only x coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withonlyxcoord")
MobiDevice("PhoneLookup").MobiEdit("Username").Click


' Step 24:  Execute DblClick  with only y coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  with only y coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withonlyycoord")
MobiDevice("PhoneLookup").MobiEdit("Username").Click


' Step 25:  Execute DblClick  with valid coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "DblClick : Execute DblClick  with valid coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
blnFlag = VerifyDblClick(objMobiEdit2, "withvalidvalues")
'MobiDevice("PhoneLookup").MobiEdit("Username").Click


' Step 26:  Execute DblClick  with negative coordinates
'#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "DblClick : Execute DblClick  with negative coordinates on MobiEdit."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute DblClick on MobiEdit." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed for negative values"
'blnFlag = VerifyDblClick(objMobiEdit2, "withnegativecoords")
'

' Step 27:  Execute Exist  when object is visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Exist : Execute Exist when object is visible on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
blnFlag = VerifyExist(objMobEdit, True, 5)
Window("regexpwndtitle:=deviceViewer").Type micReturn
' Step 28:  Execute Exist  when object is not visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Exist : Execute Exist when object is not visible on MobiEdit."
Login "mobilelabs","demo"
Window("regexpwndtitle:=deviceViewer").Type micReturn
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
blnFlag = VerifyExist(objMobEdit, False, 5)
LogOut

' Step 29:  Execute GetROProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "GetROProperty : Execute GetROProperty on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetROProperty on MobiEdit." & VBNewLine
arrProperty = Array("visible" , "nativeclass")
arrPropertyValue = Array(True ,"UITextField")
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
blnFlag = VerifyGetROProperty(objMobEdit, arrProperty, arrPropertyValue)

' Step 30:  Execute GetTOProperties 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "GetTOProperties : Execute GetTOProperties on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperties on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("id")
blnFlag = VerifyGetTOProperties(objMobEdit, arrProps)

' Step 31:  Execute GetTOProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "GetTOProperty : Execute GetTOProperty on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperty on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrPropValues = Array("2")
blnFlag = VerifyGetTOProperty(objMobEdit, arrProps, arrPropValues)

' Step 32:  Execute GetVisibleText  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "GetVisibleText : Execute GetVisibleText without coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
MobiDevice("PhoneLookup").MobiEdit("Username").Set "TestingOCR@123"
Window("regexpwndtitle:=deviceViewer").Type micReturn
blnFlag = VerifyGetVisibleText(MobiDevice("PhoneLookup").MobiEdit("Username"),False)

 'Step 33:  Execute GetVisibleText  with coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "GetVisibleText : Execute GetVisibleText with coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetVisibleText on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
MobiDevice("PhoneLookup").MobiEdit("Username").Set "TestingOCR@123"
Window("regexpwndtitle:=deviceViewer").Type micReturn
blnFlag = VerifyGetVisibleText(MobiDevice("PhoneLookup").MobiEdit("Username"),True)

' Step 34:  Execute LongClick  with valid lapse and without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and without coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "validlapsewithoutcoords")

' Step 35:  Execute LongClick  with valid lapse and with zero coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and with zero coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "validlapsezerocoords")

' Step 36:  Execute LongClick   with valid lapse  with boundary coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and with zero coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "validlapseboundarycoords")

' Step 37:  Execute LongClick  with valid lapse and with random coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and with random coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "validlapsezerocoords")

' Step 38:  Execute LongClick  with valid lapse and with x coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and with x coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "validlapsexcoords")

' Step 39:  Execute LongClick  with valid lapse and with y coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and with y coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "validlapseycoords")

' Step 40:  Execute LongClick  with valid lapse and with valid coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "LongClick : Execute LongClick with valid lapse and with valid coordinates on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
blnFlag = VerifyLongClick(objMobEdit , "withvalidlapsevalidvalue")

' Step 41:  Execute LongClick  with valid lapse and with negative coordinates
'#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "LongClick : Execute LongClick with valid lapse and with negative coordinates on MobiEdit."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute LongClick on MobiEdit." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed"
'blnFlag = VerifyLongClick(objMobEdit , "validlapsenegativecoords")
'
' Step 42:  Execute LongClick  with invalid lapse 
'#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "LongClick : Execute LongClick with invalid lapse on MobiEdit."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute LongClick on MobiEdit." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed"
'blnFlag = VerifyLongClick(objMobEdit , "withinvalidlapsetime")
'
MobiDevice("PhoneLookup").Type vbCr
' Step 43:  Execute RefreshObject
'#######################################################
'
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "RefreshObject : Execute RefreshObject on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RefreshObject  on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnFlag = VerifyRefreshObject(objMobEdit)

' Step 44:  Execute Set  with alphanumeric string
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Set : Execute Set with alphanumeric string on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Set should set the passed in text correctly."
blnFlag = VerifySet(MobiDevice("PhoneLookup").MobiEdit("Username"), "mobilelabs",null)

' Step 45:  Execute Set  with special characters 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Set : Execute Set with special characters  on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Set on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Set should set the passed in text correctly."
blnFlag = VerifySet(MobiDevice("PhoneLookup").MobiEdit("Username"), "`@#$%^&*()_+={}",null)

' Step 46:  Execute SetTOProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "SetTOProperty : Execute SetTOProperty  on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute SetTOProperty on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
arrProps = Array("id", "font")
blnFlag = VerifySetTOProperty(objMobEdit, arrProps)

' Step 47:  Execute TOString 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "TOString : Execute TOString  on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute TOString on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnFlag = VerifyTOString(objMobEdit)

' Step 48:  Execute WaitProperty  when object is visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "WaitProperty : Execute WaitProperty when object is visible on MobiEdit."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
MobiDevice("PhoneLookup").MobiEdit("Username").Set "WaitProperty"
blnFlag = VerifyWaitProperty(MobiDevice("PhoneLookup").MobiEdit("Username"), "value", "WaitProperty", 5000, True)

' Step 49:  Execute WaitProperty  when object is not visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "WaitProperty : Execute WaitProperty when object is not visible on MobiEdit."


GoToScreeniOS "Controls"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
'MobiDevice("PhoneLookup").MobiEdit("Username").Set "WaitProperty"
blnFlag = VerifyWaitProperty(MobiDevice("PhoneLookup").MobiEdit("Username"), "value", "WaitProperty", 5000, False)


' Step 6:  Execute CheckProperty when object is not visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "CheckProperty : Execute CheckProperty when object is not visible"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
'MobiDevice("PhoneLookup").MobiEdit("Username").Set "Go@Mobile"
blnFlag = VerifyCheckProperty(objMobEdit, "visible", False , 5000, False)
LogOut

'Reset the correct username
MobiDevice("PhoneLookup").MobiEdit("Username").Set "mobilelabs"

'Enter password
MobiDevice("PhoneLookup").MobiEdit("Password").Click
MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
MobiDevice("PhoneLookup").Type vbCr
Window("regexpwndtitle:=deviceViewer").Activate
Window("regexpwndtitle:=deviceViewer").Type micReturn

'*********************************************************************************************************************

EndTestIteration





