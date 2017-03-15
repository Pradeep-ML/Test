
'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiDevice object methods
' Test Description: Execute all methods for MobiDevice in Controls of PhoneLookup app.
' Expected Result: All MobiDevice methods should work correctly. The methods are: CaptureBitmap , Check , CheckProperty , ChildObjects , Click
' GetROProperty , GetTOProperties , GetTOProperty ,  'RefreshObject  , SetTOProperty , ToString , WaitProperty , Exist
'Scroll , Swipe , GetTextLocation , LongClick , DblClick , GetVisibleText

' Steps:
'Step1: Navigate to control list in PhoneLookUp
'Step2: Execute CaptureBitmap
'Step3: Execute CheckProperty
'Step4: Execute Click without coordinates
'Step5: Execute Click with coordinates
'Step6: Execute Click with boundary coordinates
'Step7: Execute Click with zero coordinates
'Step8: Execute Exist
'Step9: Execute GetTOProperties
'Step10: Execute GetTOProperty
'Step11: Execute GetTextLocation
'Step12: Execute SetTOProperty
'Step13: Execute GetVisibleText  without coordinates
'Step14: Execute GetVisibleText  with coordinates
'Step15: Execute LongClick  without coordinates
'Step16: Execute LongClick  with coordinates
'Step17: Execute LongClick  with boundary coordinates
'Step18: Execute LongClick  with zero coordinates
'Step19: Execute Scroll eBOTTOM
'Step20: Execute Scroll  eTOP
'Step21: Execute ToString
'Step22: Execute WaitProperty
'Step23: Execute Swipe eDOWN with all input parameters
'Step24:  Execute Swipe eUP with all input parameters
'Step25:  Execute Restore
'Step26:  Execute ChildObjects
'Step27:  Execute RefreshObject
'Step28:  Execute DblClick without coordinates
'Step29:  Execute DblClick with coordinates
'Step20:  Execute DblClick with boundary coordinates
'Step31:  Execute DblClick with zero coordinates
'Step32:  Execute GetROProperty
'Step33: Execute Activate
'Step34: Execute Minimize
'Step35: Execute Type
'Step36: ButtonPress

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
intStep = 0
Environment("intStepNo") = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
'#######################################################

'#######################################################
'Initial Setup

'Set object for MobiElement
Set objMobiDevice = MobiDevice("PhoneLookup")

'Step1: Navigate to control list in PhoneLookUp
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Navigate to MobiList (Control) page of PhoneLookUp App."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
'"Navigate to PickerView page of PhoneLookUp App." & VBNewLine
Environment("ExpectedResult") = "User should be navigated to UIScrollView Page"
If Not MobiDevice("PhoneLookup").MobiList("lstControls").Exist(1) Then
	LogOut
	'Login and navigate to Controls page
	blnFlag = LoginAndNavigateToControlsPage("", objMobiDevice)
Else
	'Bring the controls list to base state
	MobiDevice("PhoneLookup").MobiList("lstControls").Scroll eTOP	
	blnFlag = True
End If

' Step 2:  Execute CaptureBitmap with .png extension
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png extention on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file with .png extention to the defined location."
blnFlag = VerifyCaptureBitmap(objMobiDevice,"png")

' Step 3:  Execute CaptureBitmap with .bmp extension
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .bmp extention on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file with .bmp extention to the defined location."
blnFlag = VerifyCaptureBitmap(objMobiDevice,"bmp")

' Step 4:  Execute CaptureBitmap with .bmp extension already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .bmp extention already exist on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed"
blnFlag = VerifyCaptureBitmap(objMobiDevice,"override_bmp")

' Step 5:  Execute CaptureBitmap with .png extension already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png extention already exist on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed"
blnFlag = VerifyCaptureBitmap(objMobiDevice,"override_png")

' Step 6:  Execute CheckProperty when object exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty when object exist on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
blnFlag = VerifyCheckProperty(objMobiDevice, "name",objMobiDevice.GetROProperty("name"), 5000, True)


' Step 7:  Execute ChildObjects recursively
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ChildObjects on MobiDevice recursively."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiList." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiDevice,"recursive",50)

' Step 7:  Execute ChildObjects non-recursively
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ChildObjects on MobiDevice non-recursively."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiList." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
blnFlag = VerifyChildObjects(objMobiDevice,"nonrecursive",5)

' Step 8:  Execute Click without coordinates
'#######################################################
GoToScreeniOS "controls"
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click without coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withoutcoords")

' Step 9:  Execute Click with random coordinates 
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with random coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withrandomcoords")


' Step 10:  Execute Click with boundary coordinates at Top-Left corner 
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with boundary coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withboundarycoordsTopLeft")

' Step 10:  Execute Click with boundary coordinates at Top-Right corner 
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with boundary coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withboundarycoordsTopRight")

' Step 10:  Execute Click with boundary coordinates at Bottom-Left corner 
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with boundary coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withboundarycoordsBottomLeft")

' Step 10:  Execute Click with boundary coordinates at Bottom-Right corner 
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with boundary coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withboundarycoordsBottomRight")

' Step 10:  Execute Click with x co-ordinates
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with boundary coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withxvalue")

' Step 10:  Execute Click with y co-ordinates
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with boundary coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withyvalue")


' Step 11:  Execute Click with zero coordinates 
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with zero coordinates on MobiDevice"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiDevice, "withzerovalues")


' Step 12:  Execute Exist  when object is visible
'#######################################################
GoToScreeniOS "controls"

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist  when object is visible on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiList." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
blnFlag = VerifyExist(objMobiDevice, True, 5)


' Step 13:  Execute GetTOProperties 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperties on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("name")
blnFlag = VerifyGetTOProperties(objMobiDevice,  arrProps)


' Step 14:  Execute GetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("name")
arrPropsValue = Array("PhoneLookup")
blnFlag = VerifyGetTOProperty(objMobiDevice, arrProps, arrPropsValue)


' Step 15:  Execute RefreshObject 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RefreshObject  on MobiList." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnFlag = VerifyRefreshObject(objMobiDevice )



' Step 16:  Execute SetTOProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetTOProperty on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute SetTOProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnFlag = VerifySetTOProperty(objMobiDevice, arrProps)


' Step 17:  Execute Swipe eDOWN 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eDOWN  on MobiDevice."
'Bringing list to Base state before swipe
MobiDevice("PhoneLookup").MobiList("lstControls").Scroll eTOP

Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice , eDOWN , ,30,,ObjAfterSwipe)



' Step 18:  Execute Swipe eUP 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP  on MobiDevice."

Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP,,20,, ObjAfterSwipe)


' Step 19:  Execute Swipe eDOWN  and velocity eFAST
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eFAST  on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eFAST,30,, ObjAfterSwipe)


' Step 20:  Execute Swipe eUP and velocity eFAST
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eFAST on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST,20,, ObjAfterSwipe)


' Step 21:  Execute Swipe eDOWN  and velocity eMEDIUM
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eMEDIUM  on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eMEDIUM,40,, ObjAfterSwipe)


' Step 22:  Execute Swipe eUP and velocity eMEDIUM
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eMEDIUM on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eMEDIUM,20,, ObjAfterSwipe)


' Step 23:  Execute Swipe eDOWN  and velocity eSLOW
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eSLOW  on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eSLOW,30,, ObjAfterSwipe)


' Step 24:  Execute Swipe eUP and velocity eSLOW
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eSLOW on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eSLOW,20,, ObjAfterSwipe)


' Step 25:  Execute Swipe eDOWN  and velocity eSLOW and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eSLOW and valid starting percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eSLOW, 35, , ObjAfterSwipe)


' Step 26:  Execute Swipe eUP and velocity eSLOW and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eSLOW and valid starting percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eSLOW, 20, , ObjAfterSwipe)


' Step 27:  Execute Swipe eDOWN  and velocity eMEDIUM and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eMEDIUM and valid starting percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eMEDIUM, 30, , ObjAfterSwipe)


' Step 28:  Execute Swipe eUP and velocity eMEDIUM and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eMEDIUM and valid starting percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eMEDIUM, 20, , ObjAfterSwipe)


' Step 29:  Execute Swipe eDOWN  and velocity eFAST and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eFAST and valid starting percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eFAST, 30, , ObjAfterSwipe)


' Step 30:  Execute Swipe eUP and velocity eFAST and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eFAST and valid starting percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, 20, , ObjAfterSwipe)


' Step 31:  Execute Swipe without parameter
'#######################################################
''intStep = intStep+1
''Environment("StepName") = "Step" & intStep
''Environment("Description") = "Execute Swipe without parameter on MobiDevice."
''Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlTop")
'''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'''"Execute Swipe on MobiList." & VBNewLine
''Environment("ExpectedResult") = "Error message should be displayed."
''blnFlag = VerifySwipe(objMobiDevice , , , , , ObjAfterSwipe)

'
'' Step 52:  Execute Swipe with valid direction and invalid velocity
''#######################################################
''intStep = intStep+1
''Environment("StepName") = "Step" & intStep
''Environment("Description") = "Execute Swipe with valid direction and invalid velocity on MobiDevice."
''Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'''"Execute Swipe on MobiList." & VBNewLine
''Environment("ExpectedResult") = "Error message should be displayed."
''blnFlag = VerifySwipe(objMobiDevice , eDOWN, eABC, , , ObjAfterSwipe)
''
'
' Step 32:  Execute Swipe eDOWN  and velocity eFAST and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eFAST and valid ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eFAST, , 80, ObjAfterSwipe)


' Step 33:  Execute Swipe eUP and velocity eFAST and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eFAST and valid ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, , 80, ObjAfterSwipe)


' Step 34:  Execute Swipe eDOWN  and velocity eMEDIUM and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eMEDIUM and valid ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eMEDIUM, , 80, ObjAfterSwipe)


' Step 35:  Execute Swipe eUP and velocity eMEDIUM and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eMEDIUM and valid ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eMEDIUM, , 80, ObjAfterSwipe)


' Step 36:  Execute Swipe eDOWN  and velocity eSLOW and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eSLOW and valid ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eSLOW, , 80, ObjAfterSwipe)


' Step 37:  Execute Swipe eUP and velocity eSLOW and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eSLOW and valid ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eSLOW, , 80, ObjAfterSwipe)


' Step 38:  Execute Swipe eDOWN  and velocity eSLOW and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eSLOW and valid starting & ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eSLOW, 30, 80, ObjAfterSwipe)


' Step 39:  Execute Swipe eUP and velocity eSLOW and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eSLOW and valid starting & ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UILabel")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eSLOW, 20, 80, ObjAfterSwipe)


' Step 40:  Execute Swipe eDOWN  and velocity eMEDIUM and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eMEDIUM and valid starting & ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eMEDIUM, 30, 90, ObjAfterSwipe)


' Step 41:  Execute Swipe eUP and velocity eMEDIUM and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eMEDIUM and valid starting & ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UILabel")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eMEDIUM, 20, 80, ObjAfterSwipe)


' Step 42:  Execute Swipe eDOWN  and velocity eFAST and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and velocity eFAST and valid starting & ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIWebView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eDOWN,eFAST, 30, 80, ObjAfterSwipe)


' Step 43:  Execute Swipe eUP and velocity eFAST and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and velocity eFAST and valid starting & ending percentage on MobiDevice."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UILabel")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a MobiDevice"
blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, 20, 80, ObjAfterSwipe)


' Step 44:  Execute Swipe with valid direction and valid velocity and invalid start percentage
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe with valid direction and valid velocity and invalid start percentage on MobiDevice."
'Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error Message should be displayed"
'blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, 10.57, , ObjAfterSwipe)
'

'' Step 66:  Execute Swipe with valid direction and valid velocity and invalid end percentage
''#######################################################
''intStep = intStep+1
''Environment("StepName") = "Step" & intStep
''Environment("Description") = "Execute Swipe with valid direction and valid velocity and invalid end percentage on MobiDevice."
''Set ObjAfterSwipe =MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'''"Execute Swipe on MobiList." & VBNewLine
''Environment("ExpectedResult") = "Error Message should be displayed"
''blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, , 10.57, ObjAfterSwipe)
''
'
'' Step 67:  Execute Swipe with valid direction and valid velocity and negative start percentage
''#######################################################
''intStep = intStep+1
''Environment("StepName") = "Step" & intStep
''Environment("Description") = "Execute Swipe with valid direction and valid velocity and negative start percentage on MobiDevice."
''Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'''"Execute Swipe on MobiList." & VBNewLine
''Environment("ExpectedResult") = "Error Message should be displayed"
''blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, -10, , ObjAfterSwipe)
'
'
'' Step 68:  Execute Swipe with valid direction and valid velocity and negative end percentage
''#######################################################
''intStep = intStep+1
''Environment("StepName") = "Step" & intStep
''Environment("Description") = "Execute Swipe with valid direction and valid velocity and negative end percentage on MobiDevice."
''Set ObjAfterSwipe =MobiDevice("PhoneLookup").MobiElement("ADBannerView")
'''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'''"Execute Swipe on MobiList." & VBNewLine
''Environment("ExpectedResult") = "Error Message should be displayed"
''blnFlag = VerifySwipe(objMobiDevice ,eUP ,eFAST, , -10, ObjAfterSwipe)
'

' Step 44:  Execute ToString
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute TOString on MobiList." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnFlag = VerifyTOString(objMobiDevice)


' Step 45:  Execute WaitProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is visible on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnFlag = VerifyWaitProperty(objMobiDevice, "name", "PhoneLookup", 5000, True)


' Step 46:  Execute GetROProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiDevice. "
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetROProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
arrProperty = Array("name")
arrPropertyValue = Array("PhoneLookup")
blnFlag = VerifyGetROProperty(objMobiDevice, arrProperty, arrPropertyValue)


'Step 47:  Execute Restore
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Restore  on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Restore  on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Restore should work"
blnFlag = VerfiyRestore(objMobiDevice)


'Step 48:  Execute RefreshObject
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject  on  MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RefreshObject  on  MobiDevice." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnFlag = VerifyRefreshObject(objMobiDevice)


'Step 49: Execute Activate
'######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Activate on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Activate on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Activates the MobiDevice window."
blnFlag = verifyActivate(objMobiDevice)


'Step 50: Execute Minimize
'######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Minimize on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Minimize on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Minimize the mobidevice window into an icon"
blnFlag = VerifyMinimize(objMobiDevice)

Step 51: Execute GetTextLocation
'######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTextLocation on MobiDevice."
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTextLocation  on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "GetTextLocation should return window area for the specified string"
blnFlag = VerifyGetTextLocation(objMobiDevice , "UIButton" , True)


'Step 52: Execute GetVisibleText  without coordinates
'######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetvisibleText  without coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetvisibleText on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "GetvisibleText  should return text within specified area."
blnFlag = VerifyGetVisibleText(objMobiDevice , False)

'Step 53: Execute GetVisibleText  with coordinates
'######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetvisibleText  with coordinates on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetvisibleText on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "GetvisibleText  should return text within specified area."
blnFlag = VerifyGetVisibleText(objMobiDevice , True)


' Step 54:  Execute LongClick With Invalid Valid Lapse
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute LongClick With negative co-ordinates"
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "LongClick should throw an error"
'
'blnFlag = VerifyLongClick(objMobiDevice ,"validlapsenegativecoords") 

' Step19:  Execute LongClick With Invalid Valid Lapse
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute LongClick With inValid Lapse"
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "LongClick should throw an error"
'
'blnFlag = VerifyLongClick(objMobiDevice ,"withinvalidlapsetime") 

' Step 54:  Execute LongClick with VAlid Lapse without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick with Valid Lapse and  valid x and y co-ordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

blnFlag = VerifyLongClick(objMobiDevice  , "withvalidlapsevalidvalue")

GoToScreeniOS "controls"
 
' Step 55:  Execute LongClick with VAlid Lapse without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick with VAlid Lapse without coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

blnFlag = VerifyLongClick(objMobiDevice  , "validlapsewithoutcoords")
GoToScreeniOS "controls"

' Step 56:  Execute LongClick With Valid Lapse At 0,0 Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse At 0,0 Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

blnFlag = VerifyLongClick(objMobiDevice  , "validlapsezerocoords")
GoToScreeniOS "controls"
'
' Step 57:  Execute LongClick With Valid Lapse At Boundary Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse At Boundary Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

blnFlag = VerifyLongClick(objMobiDevice  , "validlapseboundarycoords")
GoToScreeniOS "controls"


' Step 58:  Execute DblClick  with Valid x & y  coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With valid X & Y coords"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withvalidvalues")
GoToScreeniOS "controls"

' Step 59:  Execute DblClick  withoutcoords  
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At without Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withoutcoords")

GoToScreeniOS "controls"

' Step 60:  Execute DblClick  withboundarycoords
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At Boundary Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withboundarycoords")

GoToScreeniOS "controls"

' Step 61:  Execute DblClick  withrandomcoords  
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At random Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withrandomcoords")
GoToScreeniOS "controls"

' Step 62:  Execute DblClick  withzercoords
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At zero Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withzercoords")

GoToScreeniOS "controls"

' Step 63:  Execute DblClick  withnegativecoords
'#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute DblClick With Valid Lapse At negative Coordinates"
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "DblClick should throw an error"
'blnFlag = VerifyDblClick(objMobiDevice  , "withnegativecoords")


' Step 64:  Execute DblClick   withonlyxcoord
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick At x Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withonlyxcoord")

GoToScreeniOS "controls"

' Step 65:  Execute DblClick   withonlyycoord
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At y Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
blnFlag = VerifyDblClick(objMobiDevice  , "withonlyycoord")

GoToScreeniOS "controls"

LogOut

'Step 66: Execute Type
'######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Type on MobiDevice."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Type on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Navigate to Login Page to test Type method"

blnFlag = VerifyType(objMobiDevice , "test the type method")


' Step 116   Execute Scale with Blank Value
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with Blank Value on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with Blank Value on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with Blank Value should throw an error"
strResult = VerifyScale(objMobiDevice,"")

' Step 116   Execute Scale with a string Value
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with String Value on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with String Value on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with string Value should throw an error"
strResult = VerifyScale(objMobiDevice,"Hello")

' Step 116   Execute Scale with a float/Double Value
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with float/Double Value on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with float/Double Value on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with float/Double Value should throw an error"
strResult = VerifyScale(objMobiDevice,56.9)

' Step 116   Execute Scale with less than 25 Value
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with less than 25 Value on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with less than 25 Value on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with less than 25 Value should throw an error"
strResult = VerifyScale(objMobiDevice,10)

' Step 116   Execute Scale with greater than 100 Value
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with greater than 100 Value on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with greater than 100 Value on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with greater than 100 Value should throw an error"
strResult = VerifyScale(objMobiDevice,500)

' Step 116   Execute Scale with valid Value 25
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with valid Value 25 on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with valid Value 25 on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with valid Value 25 should change the scale to 25"
strResult = VerifyScale(objMobiDevice,25)

' Step 116   Execute Scale with valid Value 80
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with valid Value 80 on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with valid Value 80 on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with valid Value 80 should change the scale to 80"
strResult = VerifyScale(objMobiDevice,80)

' Step 116   Execute Scale with valid Value 100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scale method with valid Value 100 on MobiDevice"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scale method with valid Value 100 on MobiDevice." & VBNewLine
Environment("ExpectedResult") = "Scale method with valid Value 100 should change the scale to 100"
strResult = VerifyScale(objMobiDevice,100)
'###############################################################
EndTestIteration












