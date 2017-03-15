'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiList object
' Test Description: Execute all methods for MobiListon on Controls List
' Expected Result: All MobiList methods should work correctly. The methods are: CaptureBitmap, CheckProperty, ChildObjects, 
' Click, Exist, GetItem, GetROProperty, GetTOProperties, GetTOProperty, RefreshObject, RowCount, 
' Scroll, Select, SetTOProperty, Swipe, TOString, WaitProperty

' Steps:
' Step1: Navigate to MobiList (Control) page of PhoneLookUp App.
'Step2:  Execute CaptureBitmap 
'Step3:  Execute CheckProperty 
'Step4:  Execute ChildObjects 
' Step5:  Execute Click without coordinates
' Step6:  Execute Click with random coordinates 
' Step7:  Execute Click with boundary coordinates
' Step8:  Execute Click with zero coordinates
' Step9:  Execute Exist 
' Step10:  Execute GetItem  for first item in the list
' Step11:  Execute GetItem  for last item in the list
' Step12:  Execute GetROProperty
' Step13:  Execute GetTOProperties 
' Step14:  Execute GetTOProperty 
' Step15:  Execute RefreshObject 
' Step16:  Execute RowCount 
' Step17:  Execute Scroll  eTOP 
' Step18:  Execute Scroll eBOTTOM 
' Step19:  Execute Select with String  as input
' Step20:  Execute Select with index as input
' Step21:  Execute Select with index  in quotes as input
' Step22:  Execute Select  for last item with String  as input
' Step23:  Execute Select  for last item with index as input
' Step24:  Execute Select  for last item with index  in quotes as input
' Step25:  Execute SetTOProperty
' Step26:  Execute Swipe eDOWN with all input parameters
' Step27:  Execute Swipe eUP with all input parameters
' Step28:  Execute Swipe eDOWN without start and end percentage
' Step29:  Execute Swipe eUP without start and end percentage
' Step30:  Execute ToString
' Step31:  Execute WaitProperty


'##########################################################################################################

'#######################################################
'Declare Variables
Dim strStepsToReproduce
Dim strStepName
Dim intStep
Dim blnResult
Dim strTestName
CreateReportTemplate
''#######################################################

'#######################################################
'Initializations
intStep = 0
Environment("intStepNo") = 0
Environment("Component") = "PhoneLookup_ObjectBased"
Environment("StepsToReproduce") = ""
'#######################################################

' Step1: Navigate to MobiList (Control) page of PhoneLookUp App.
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

''Set object for List Controls
Set objMobList = MobiDevice("PhoneLookup").MobiList("lstControls")


	Logout
	Environment("Description") = "Navigate to MobiList (Control) page of PhoneLookUp App."
'	Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
'	"Navigate to MobiList (Control) page of PhoneLookUp App." & VBNewLine
	Environment("ExpectedResult") = "User should be navigated to MobiList (Control) Page"
	
blnFlag  = LoginAndNavigateToControlsPage("", objMobList)

' Step 2:  Execute CaptureBitmap with .png extension
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png extention on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file with .png extention to the defined location."
blnFlag = VerifyCaptureBitmap(objMobList,"png")

' Step 3:  Execute CaptureBitmap with .bmp extension
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .bmp extention on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file with .bmp extention to the defined location."
blnFlag = VerifyCaptureBitmap(objMobList,"bmp")

' Step 4:  Execute CaptureBitmap with .bmp extension already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .bmp extention already exist on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed"
blnFlag = VerifyCaptureBitmap(objMobList,"override_bmp")

' Step 5:  Execute CaptureBitmap with .png extension already exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png extention already exist on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CaptureBitmap on MobiList." & VBNewLine
Environment("ExpectedResult") = "Error message should be displayed"
blnFlag = VerifyCaptureBitmap(objMobList,"override_png")

' Step 6:  Execute CheckProperty when object exist
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty when object exist on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
blnFlag = VerifyCheckProperty(objMobList, "visible", True , 5000, True)


'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobList, "nonrecursive" , 24 )

'Step 7 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobList, "recursive" , 34 )

'Step 20:  Execute Click  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click on MobList without coordinates."
Environment("ExpectedResult") = "Click should work correctly without co-ordinates."
blnStepRC = VerifyClick(objMobList, "withoutcoords")
'Open alternative URL

GoToScreeniOS  "controls"
'Step 21:  Execute Click with  random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList for random co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with random co-ordinates."
blnStepRC = VerifyClick(objMobList, "withrandomcoords")
'Open alternative URL

GoToScreeniOS  "controls"

'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."
blnStepRC = VerifyClick(objMobList, "withboundarycoordsTopLeft")
'Open alternative URL

GoToScreeniOS  "controls"
'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."
blnStepRC = VerifyClick(objMobList, "withboundarycoordsTopRight")
'Open alternative URL

GoToScreeniOS  "controls"
'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."
blnStepRC = VerifyClick(objMobList, "withboundarycoordsBottomLeft")
'Open alternative URL

GoToScreeniOS  "controls"

'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."
blnStepRC = VerifyClick(objMobList, "withboundarycoordsBottomRight")
'Open alternative URL

GoToScreeniOS  "controls"

'Step 26:  Execute Click with x co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList with x co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with x co-ordinates."
blnStepRC = VerifyClick(objMobList, "withxvalue")
'Open alternative URL

GoToScreeniOS  "controls"

'Step 27:  Execute Click with y co-ordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobList with y co-ordinates."
Environment("ExpectedResult") = "Click should work correctly with y co-ordinates."
blnStepRC = VerifyClick(objMobList, "withyvalue")


'Open alternative URL

GoToScreeniOS  "controls"

' Step 12:  Execute Click with valid x and y coordinates 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click with valid x and y coordinates on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiList." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobList, "withvalidvalue")

GoToScreeniOS  "controls"

'' Step 12:  Execute Click with negative coordinates 
''#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Click with negative coordinates on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Click on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed"
'blnFlag = VerifyClick(objMobList, "withnegativecoords")


' Step 13:  Execute Exist  when object is visible
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist  when object is visible on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiList." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
blnFlag = VerifyExist(objMobList, True, 5)


' Step 14:  Execute GetItem  for first item in the list with index
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetItem with index on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetItem on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetItem should get the correct run-time value for the specifed index location."
blnFlag = VerifyGetItem(objMobList,0,0,"CustomTextController","withindexonly")

'' Step 15:  Execute GetItem  for first item in the list with index as string
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute GetItem with index as string on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute GetItem on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifyGetItem(objMobList,"0",0,"MKMapView","withindexonly")

'' Step 16:  Execute GetItem  for first item in the list with index out of range
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute GetItem  with index out of range on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute GetItem on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifyGetItem(objMobList, 24 , 0, "MKMapView", "withindexonly")

'' Step 17:  Execute GetItem  for first item in the list without parameter
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute GetItem without parameter on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute GetItem on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifyGetItem(objMobList, 0 , 0, "MKMapView", "withoutparameter") 

'' Step 18:  Execute GetItem  for first item in the list with negative index
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute GetItem  with negative index on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute GetItem on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifyGetItem(objMobList, -1 , 0, "MKMapView", "withindexonly")

' Step 19:  Execute GetItem  for last item in the list
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetItem  for last item in the list on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetItem on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetItem should get the correct run-time value for the specifed index location."
blnFlag = VerifyGetItem(objMobList, objMobList.RowCount -1, 0 , "UIWebView", "withindexonly")


' Step 20:  Execute GetTOProperties 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperties on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("listtype", "allowmultipleselection")
blnFlag = VerifyGetTOProperties(objMobList,  arrProps)


' Step 21:  Execute GetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTOProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("listtype", "allowmultipleselection")
arrPropsValue = Array("0","False")
blnFlag = VerifyGetTOProperty(objMobList, arrProps, arrPropsValue)


' Step 22:  Execute RefreshObject 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RefreshObject on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RefreshObject  on MobiList." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
blnFlag = VerifyRefreshObject(objMobList )


' Step 23:  Execute RowCount 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RowCount on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute RowCount  on MobiList." & VBNewLine
Environment("ExpectedResult") = "RowCount represents number of rows contained in a list"
blnFlag = VerifyRowCount(objMobList , 23, "")


' Step 24:  Execute Scroll  eTOP 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll  eTOP  on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Scroll up on MobiList." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards up."

objMobList.Scroll eBOTTOM
Set objListControlTop =MobiDevice("PhoneLookup").MobiElement("UILabel")
blnFlag = VerifyScroll(objMobList, "top", objListControlTop)

' Step 25:  Execute Scroll eBOTTOM 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll  eBOTTOM  on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Scroll down on MobiList." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards down."
Set objListControlBottom = MobiDevice("PhoneLookup").MobiElement("UIWebView")
blnFlag = VerifyScroll(objMobList, "bottom", objListControlBottom)


'' Step 26:  Execute Scroll without parameter 
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Scroll  without parameter  on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Scroll down on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed"
'Set objListControlBottom =  MobiDevice("PhoneLookup").MobiWebView("WebView")
'blnFlag = VerifyScroll(objMobList, "withoutparameter", objListControlBottom)


' Step 35:  Execute SetTOProperty
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetTOProperty on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute SetTOProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
blnFlag = VerifySetTOProperty(objMobList, arrProps)

MobiDevice("PhoneLookup").MobiList("List").Scroll eTOP

' Step 37:  Execute Swipe eDOWN 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eDOWN  on MobiList."
'Bringing list to Base state before swipe
objMobList.Scroll eTOP

Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlBottom")

'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList , eDOWN, , 20,80,ObjAfterSwipe)


' Step 38:  Execute Swipe eUP 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP  and velocity eSLOW on MobiList."

Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP,eSLOW,, ,ObjAfterSwipe)


' Step 39:  Execute Swipe eDOWN  and valocity eFAST
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eFAST  on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlBottom")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eFAST,20,70, ObjAfterSwipe)


' Step 40:  Execute Swipe eUP and valocity eFAST
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eFAST on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eFAST,20,70, ObjAfterSwipe)


' Step 41:  Execute Swipe eDOWN  and valocity eMEDIUM
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eMEDIUM  on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlBottom")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eMEDIUM,20,70, ObjAfterSwipe)


' Step 42:  Execute Swipe eUP and valocity eMEDIUM
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eMEDIUM on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eMEDIUM,20,70, ObjAfterSwipe)


' Step 43:  Execute Swipe eDOWN  and valocity eSLOW
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eSLOW  on MobiList."
objMobList.Scroll eTOP
Wait 3
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UITabBar")

'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
MobiDevice("PhoneLookup").MobiList("lstControls").Swipe eDOWN , eSLOW , 20 , 70 
blnFlag = VerifySwipe(objMobList ,eDOWN,eSLOW,20,70, ObjAfterSwipe)


' Step 44:  Execute Swipe eUP and valocity eSLOW
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eSLOW on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eSLOW,20 ,70, ObjAfterSwipe)


' Step 45:  Execute Swipe eDOWN  and valocity eSLOW and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eSLOW and valid starting percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlBottom")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eSLOW, 30, 80, ObjAfterSwipe)


' Step 46:  Execute Swipe eUP and valocity eSLOW and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eSLOW and valid starting percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eSLOW, 30,80 , ObjAfterSwipe)


' Step 47:  Execute Swipe eDOWN  and valocity eMEDIUM and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eMEDIUM and valid starting percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlBottom")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eMEDIUM, 30, 80, ObjAfterSwipe)


' Step 48:  Execute Swipe eUP and valocity eMEDIUM and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eMEDIUM and valid starting percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eMEDIUM, 30,80 , ObjAfterSwipe)


' Step 49:  Execute Swipe eDOWN  and valocity eFAST and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eFAST and valid starting percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlBottom")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eFAST, 30, 80, ObjAfterSwipe)


' Step 50:  Execute Swipe eUP and valocity eFAST and valid starting percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eFAST and valid starting percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eFAST, 30, 80, ObjAfterSwipe)


'' Step 51:  Execute Swipe without parameter
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe without parameter on MobiList."
'Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("ListControlTop")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifySwipe(objMobList , , , , , ObjAfterSwipe)


'' Step 52:  Execute Swipe with valid direction and invalid valocity
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe with valid direction and invalid valocity on MobiList."
'Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifySwipe(objMobList , eDOWN, eABC, , , ObjAfterSwipe)


' Step 53:  Execute Swipe eDOWN  and valocity eFAST and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eFAST and valid ending percentage on MobiList."

'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
objMobList.Scroll eTOP
Wait 3
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UITabBar")
blnFlag = VerifySwipe(objMobList ,eDOWN,eFAST,20 , 70, ObjAfterSwipe)


' Step 56:  Execute Swipe eUP and valocity eMEDIUM and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eMEDIUM and valid ending percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eMEDIUM, 20, 70, ObjAfterSwipe)


' Step 57:  Execute Swipe eDOWN  and valocity eSLOW and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eSLOW and valid ending percentage on MobiList."
Set ObjAfterSwipe =MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("UIScrollView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eSLOW, 20, 70, ObjAfterSwipe)


' Step 58:  Execute Swipe eUP and valocity eSLOW and valid ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eSLOW and valid ending percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eSLOW,20 , 70, ObjAfterSwipe)


' Step 59:  Execute Swipe eDOWN  and valocity eSLOW and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eSLOW and valid starting & ending percentage on MobiList."
objMobList.Scroll eTOP
Wait 3
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UITabBar")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eSLOW, 30, 70, ObjAfterSwipe)


' Step 60:  Execute Swipe eUP and valocity eSLOW and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eSLOW and valid starting & ending percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eSLOW, 30, 70, ObjAfterSwipe)


' Step 61:  Execute Swipe eDOWN  and valocity eMEDIUM and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eMEDIUM and valid starting & ending percentage on MobiList."
Set ObjAfterSwipe =  MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("UIPickerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eMEDIUM, 30, 70, ObjAfterSwipe)


' Step 62:  Execute Swipe eUP and valocity eMEDIUM and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eMEDIUM and valid starting & ending percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eMEDIUM, 30, 70, ObjAfterSwipe)


' Step 63:  Execute Swipe eDOWN  and valocity eFAST and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = " Execute Swipe eDOWN  and valocity eFAST and valid starting & ending percentage on MobiList."
Set ObjAfterSwipe =  MobiDevice("PhoneLookup").MobiList("lstControls").MobiElement("UIPickerView")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eDOWN,eFAST, 30, 70, ObjAfterSwipe)


' Step 64:  Execute Swipe eUP and valocity eFAST and valid starting & ending percentage
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eUP and valocity eFAST and valid starting & ending percentage on MobiList."
Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("UIButton")
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on a mobile list"
blnFlag = VerifySwipe(objMobList ,eUP ,eFAST, 30, 70, ObjAfterSwipe)


'' Step 65:  Execute Swipe with valid direction and valid valocity and invalid start percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe with valid direction and valid valocity and invalid start percentage on MobiList."
'Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error Message should be displayed"
'blnFlag = VerifySwipe(objMobList ,eUP ,eFAST, 10.57, , ObjAfterSwipe)
'
'
'' Step 66:  Execute Swipe with valid direction and valid valocity and invalid end percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe with valid direction and valid valocity and invalid end percentage on MobiList."
'Set ObjAfterSwipe =MobiDevice("PhoneLookup").MobiElement("ADBannerView")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error Message should be displayed"
'blnFlag = VerifySwipe(objMobList ,eUP ,eFAST, , 10.57, ObjAfterSwipe)
'

'' Step 67:  Execute Swipe with valid direction and valid valocity and negative start percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe with valid direction and valid valocity and negative end percentage on MobiList."
'Set ObjAfterSwipe = MobiDevice("PhoneLookup").MobiElement("ADBannerView")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error Message should be displayed"
'blnFlag = VerifySwipe(objMobList ,eUP ,eFAST, -10, , ObjAfterSwipe)


'' Step 68:  Execute Swipe with valid direction and valid valocity and negative end percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Swipe with valid direction and valid valocity and negative end percentage on MobiList."
'Set ObjAfterSwipe =MobiDevice("PhoneLookup").MobiElement("ADBannerView")
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Swipe on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error Message should be displayed"
'blnFlag = VerifySwipe(objMobList ,eUP ,eFAST, , -10, ObjAfterSwipe)
'

' Step 69:  Execute ToString
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute TOString on MobiList." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
blnFlag = VerifyTOString(objMobList)


' Step 70:  Execute WaitProperty when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is visible on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnFlag = VerifyWaitProperty(objMobList, "visible", "True", 5000, True)


' Step 72:  Execute GetROProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty on MobiList. "
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetROProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
arrProperty = Array("name","nativeclass" , "accessibilitylabel")
arrPropertyValue = Array("List","UITableView", "")
Set objMobList = MobiDevice("PhoneLookup").MobiList("List")
blnFlag = VerifyGetROProperty(objMobList, arrProperty, arrPropertyValue)
'#############################################################


' Step 27:  Execute Select with String  as input
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with String  as input  on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."
Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiSlider("Slider")
blnFlag = VerifySelect(objMobList, "selectstring", "UISlider", objImageAfterSelection)

GoToScreeniOS "controls"
Wait 3

' Step 28:  Execute Select with index as input
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with index  as input  on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."
Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiButton("btnGray")
blnFlag = VerifySelect(objMobList, "selectindex",  3 , objImageAfterSelection)
GoToScreeniOS "controls"
	
' Step  29:  Execute Select with index  in quotes as input
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select with index  in quotes as input on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."

Set objMobList =  MobiDevice("PhoneLookup").MobiList("lstControls")
Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiButton("btnGray")
Wait 3
blnFlag = VerifySelect(objMobList, "selectindex", "#3", objImageAfterSelection)
GoToScreeniOS "controls"


' Step 20:  Execute Select  for last item with String  as input
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select  for last item with String  as input on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."
Set objMobList = MobiDevice("PhoneLookup").MobiList("lstControls")
Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiWebView("WebView")
IntIndexLastItem = objMobList.RowCount-1
objMobList.Scroll eBOTTOM
Wait 3
blnFlag = VerifySelect(objMobList, "selectstring", objMobList.GetItem(IntIndexLastItem) , objImageAfterSelection)
GoToScreeniOS "controls"


'' Step 31:  Execute Select  with negative index
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Select  with negative index on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Select on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifySelect(objMobList, "selectnegativeindex", -1 , "")
'
'' Step 32:  Execute Select  without parameter
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Select  without parameter on MobiList."
''Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
''"Execute Select on MobiList." & VBNewLine
'Environment("ExpectedResult") = "Error message should be displayed."
'blnFlag = VerifySelect(objMobList, "withoutparameter", "" , "")
'

' Step 33:  Execute Select  for last item with index as input
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select  for last item with index as input on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."
Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiWebView("WebView")
'Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiElement("eleWebView")
blnFlag = VerifySelect(objMobList, "selectindex",  22 , objImageAfterSelection)
GoToScreeniOS "controls"



''Step 23 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") = "Verify IsOccluded  is working correctly when object  is in view without passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobList , "withoutcoords" , "notoccluded")
''#############################################################
'
''Step 24 IsOccluded  (For IsOccluded when object is in view)
''###############################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Verify IsOccluded" 
'Environment("ExpectedResult") =  "Verify IsOccluded  is working correctly when object  is in view by passing co-ordinates."
'blnResult =  VerifyIsOccluded(objMobList , "withcentervalues" , "notoccluded")
''#############################################################
'
' Step 34:  Execute Select  for last item with index  in quotes as input
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Select  for last item with index  in quotes as input on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."
Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiWebView("WebView")
'Set objImageAfterSelection = MobiDevice("PhoneLookup").MobiElement("eleWebView")
Wait 2
blnFlag = VerifySelect(objMobList, "selectindex", "#22", objImageAfterSelection)
GoToScreeniOS "controls"

'Logout
LogOut

' Step 7:  Execute CheckProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CheckProperty when object is not visible on MobiList."

'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute CheckProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
blnFlag = VerifyCheckProperty(objMobList, "visible", True , 5000, False)


' Step 13:  Execute Exist  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist  when object is not visible on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Exist on MobiList." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
blnFlag = VerifyExist(objMobList, False, 5)


' Step 71:  Execute WaitProperty when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute WaitProperty when object is not visible on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute WaitProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
blnFlag = VerifyWaitProperty(objMobList, "visible", True, 5000, False)

'###################################################################
EndTestIteration



























