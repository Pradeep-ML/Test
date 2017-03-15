
'##########################################################################################################
' Objective: Login to the PhoneLookup app and test MobiList
' Test Description: Execute all MobiList methods on Controls screen
'#######################################################
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

' Step: Navigate to Controls screen
'Expected Result: Controls screen should be displayed
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Navigate to Search screen" & VBNewLine
Environment("ExpectedResult") = "Controls screen should be displayed"

'Call function to createreporttemplare
CreateReportTemplate()

'Set object for List
Set objMobiList = MobiDevice("Phone Lookup").MobiList("List")

'Call function to navigate to Controls screen
StrResult = NavigateScreenOnPhoneLookup("Controls"  , objMobiList , "")


'*********************************************************************************************************************
' Step1:Execute CaptureBitmap with .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on mobilist." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the png file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiList , "png")

' Step2:Execute CaptureBitmap with .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with  .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on mobilist ." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the bmp file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiList , "bmp")

' Step3:Execute CaptureBitmap with .override.bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with override .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on mobilist." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw an error messge for override messagefor .bmp file."
strResult =  VerifyCaptureBitmap(objMobiList , "override_bmp")

' Step4:Execute CaptureBitmap with .override.png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on mobilist." & VBNewLine
Environment("ExpectedResult") =  "CaptureBitmap should throw an error messge for override message for .png  file."
strResult =  VerifyCaptureBitmap(objMobiList , "override_png")

' Step 5 Execute CheckProperty  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on mobilist ." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should return True."
strResult = VerifyCheckProperty(objMobiList, "visible" ,True , 5000, True)

' Step7:  Execute Exist  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on .mobilist" & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = VerifyExist(objMobiList, True, 5)

' Step9:  Execute GetTOProperties 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on mobilist ." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("micclass","allowmultipleselection","id")
strResult = VerifyGetTOProperties(objMobiList, arrProps)


' Step10  Execute GetROProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on mobilist" & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
arrProps = Array("itemscount","allowmultipleselection")
arrvalue= Array (31,"False")
strResult =VerifyGetROProperty(objMobiList, arrProps, arrvalue)

' Step11:  Execute GetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on mobilist." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("allowmultipleselection")
arrvalue= Array ("False")
strResult =VerifyGetTOProperty(objMobiList, arrProps, arrvalue)

' Step12:  Execute ToString 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ToString on mobilist." & VBNewLine
Environment("ExpectedResult") = "ToString should return the object type and class."
strResult = VerifyToString(objMobiList)

' Step13  Execute WaitProperty  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Wait property when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on mobilist" & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value"
strResult = VerifyWaitProperty(objMobiList, "visible", True , 5000 , True)

 'Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute childobject method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on mobilist." & VBNewLine

Environment("ExpectedResult") = "Return child object recursively in the application"
'blnFlag = VerifyChildObjects(objMobiList  ,"recursive",26)
blnFlag = VerifyChildObjects(objMobiList, "recursive" , 11)

 'Step 7:  Execute  ChildObjects non recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute childobject method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on mobilist." & VBNewLine

Environment("ExpectedResult") = "Return child object non recursively in the application"
'blnFlag = VerifyChildObjects(objMobiList  ,"recursive",26)
blnFlag = VerifyChildObjects(objMobiList, "nonrecursive" , 11)

' Step 7:  Execute  ChildObjects nonrecusrively
'#######################################################
'
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "ChildObjects : Execute ChildObjects on MobiList non-recursively"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute ChildObjects on MobiSlider." & VBNewLine
'Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
'blnFlag = VerifyChildObjects(objMobiList  ,"nonrecursive",1)
'
' 'Step 13:  Execute Click  with boundary coordinates at Top-Left corner
''#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiList ,"withboundarycoordsTopLeft")
GoToScreen "Controls"
'
'' Step 13:  Execute Click  with boundary coordinates at Top-Right corner
''#######################################################
'
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Right corner on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiList,"withboundarycoordsTopRight")
GoToScreen "Controls"
'
'' Step 13:  Execute Click  with boundary coordinates at Bottom-Left corner
''#######################################################
'
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Left corner on MobiList"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiList,"withboundarycoordsBottomLeft")
GoToScreen "Controls"
'
' Step 13:  Execute Click  with boundary coordinates at Bottom-Right corner
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Bottom-Right corner on MobiDropdown."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiList,"withboundarycoordsBottomRight")
GoToScreen "Controls"


'19'Execute Click   Without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method Without coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on mobilist." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiList, "withoutcoords")

GoToScreen "Controls"





'21'Execute Click with  x coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on mobilist." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiList, "withxvalue")

GoToScreen "Controls"


'22'Execute Click with  y coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on mobilist." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiList, "withyvalue")

GoToScreen "Controls"


'23'Execute Click with  Valid X & Y coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with Valis x & y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on mobilist." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiList, "withvalidvalue")

GoToScreen "Controls"


'24'Execute Click with  Random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with Random coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on mobilist." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."

strResult = VerifyClick(objMobiList, "withrandomcoords")

GoToScreen "Controls"




'Step 26 : Execute GetItem With Index 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Getitem With Index"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetItem on MobiList." & VBNewLine
Environment("ExpectedResult") = "GetItem should get the correct run-time value for the specifed index location."

strResult = VerifyGetItem(objMobiList, 8,,"GridView" , "withindexonly")

'



'Step 30 : Execute RowCount  With Blank Value
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute RowCount  With Blank Value"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RowCount  on MobiList." & VBNewLine
Environment("ExpectedResult") = "RowCount represents number of rows contained in a list"
strResult = VerifyRowCount(objMobiList , 31 , "")

'Step 35: Execute 'Select  with Item as String Case Sensitive
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute 'Select  with Item as String Case Sensitive"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."

'Set the object  that appear after select opeartion
Set objImageAfterSelection = MobiDevice("Phone Lookup").MobiButton("Changethedate")

strResult = VerifySelect(objMobiList ,"selectstring", "DatePicker" , objAfterSelection)

GoToScreen "Controls"

GoToScreen "Controls"

wait 3


'Step 36: Execute 'Select  Item as Index
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute 'Select  with Item as Index"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."

strResult = VerifySelect(objMobiList ,"selectindex", 5 , objAfterSelection)

GoToScreen "Controls"

GoToScreen "Controls"


'Step 39: Execute 'Select  with Item as String Case inSensitive
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute 'Select  with Item as String Case inSensitive"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Select on MobiList." & VBNewLine
Environment("ExpectedResult") = "Select should select the item correctly."

strResult = VerifySelect(objMobiList ,"selectstring", "DATEPicker" , objAfterSelection)

GoToScreen "Controls"


' Step 72 :Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiList." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
arrProps = Array("itemscount","allowmultipleselection","id")
strResult = VerifySetTOProperty(objMobiList, arrProps)



' Step16  Execute RefreshObject 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Refresh method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on mobilist." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiList)



' Step 32 Execute Scroll  Bottom
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll bottom"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiLIst." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards down."

Set objListControlBottom = MobiDevice("Phone Lookup").MobiElement("ZoomControls")

strResult = VerifyScroll(objMobiList, "bottom", objListControlBottom)


' Step 33:  Execute Scroll  TOP
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll Top"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiLIst." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards top."

Set objListControlTop = MobiDevice("Phone Lookup").MobiElement("AbsoluteLayout")

strResult = VerifyScroll(objMobiList, "top", objListControlTop)

'
' Step 40:  Execute Swipe down
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe edown"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a swipe edown gesture on a MobiList"

'Swipe
Set obj_up =MobiDevice("Phone Lookup").MobiElement("CheckBox")

' MobiDevice("Phone Lookup").MobiList("CheckBox")
Set obj_down =MobiDevice("Phone Lookup").MobiElement("ListView")

' MobiDevice("Phone Lookup").MobiList("ListView")

'Set obj_down = MobiDevice("Phone Lookup").MobiList("eleBottom_Swipe")
'Set obj_up = MobiDevice("Phone Lookup").MobiList("eleTop")
wait 3

'strResult = VerifySwipe(objMobiList ,eDOWN ,,,,obj_down)
'strResult = VerifySwipe(objMobiList ,eDOWN ,,,,obj_down)


MobiDevice("Phone Lookup").MobiList("List").Swipe eDOWN , eFAST , 30 , 80
wait 3
' Step 41:  Execute Swipe up
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe up"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a Swipe up gesture on a MobiList"

strResult = VerifySwipe(objMobiList , eUP , , , , obj_up)

'' Step 50:  Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction edown and velocity emedium and starting percentage 0-99 "

strResult = VerifySwipe(objMobiList , eDOWN , eMEDIUM  ,20 ,  ,obj_down)



' Step 43:  Execute  Swipe with directions as eup and velocity eFast
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity up"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity efast"

strResult = VerifySwipe(objMobiList  , eUP ,eFAST ,  , ,obj_up)

'' Step 50:  Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction edown and velocity emedium and starting percentage 0-99 "

strResult = VerifySwipe(objMobiList , eDOWN , eMEDIUM  ,30 ,  ,obj_down)


' Step 45:  Execute  Swipe with directions as eup and velocity emedium
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity emedium"
Environment("ExpectedResult") = "Simulates a gesture on a MobiList with direction eup and velocity emedium"

strResult = VerifySwipe(objMobiList  ,eUP ,eMEDIUM ,  , ,obj_up)

'' Step 48:  Execute Swipe  directions as edown and velocity eFast  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction edown and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiList , eDOWN , eFAST  ,30 ,  ,obj_down)

'
' Step 47:  Execute  Swipe with directions as eup and velocity eslow
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity eslow"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity eslow"

strResult = VerifySwipe(objMobiList  ,eUP ,eSLOW ,  , ,obj_up)





' Step 49  Execute Swipe  directions as eup and velocity eFast and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eFast  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity efast and starting percentage 0-99"

strResult = VerifySwipe(objMobiList  ,  eUP ,eFAST , 30 , ,obj_up)

'' Step 52  Execute Swipe  directions as edown and velocity eslow  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eslow and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction eslow and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiList , eDOWN , eSLOW ,30 , ,obj_down)


' Step 51  Execute Swipe  directions as eup and velocity emedium and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity emedium and starting percentage 0-99"

strResult = VerifySwipe(objMobiList  ,  eUP ,eMEDIUM , 20 ,,obj_up)

'' Step 56:  Execute Swipe  directions as edown and velocity eFast  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction edown and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiList , eDOWN , eFAST  ,20 , 80 ,obj_down)


' Step 53  Execute Swipe  directions as eup and velocity eslow and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eslow  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity eslow and starting percentage 0-99"

strResult = VerifySwipe(objMobiList  ,  eUP ,eSLOW , 20 , ,obj_up)




' Step 57  Execute Swipe  directions as eup and velocity eFast and and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eFast  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity efast and ending percentage 1-100"

strResult = VerifySwipe(objMobiList  ,  eUP ,eFAST ,  , 80,obj_up)

'' Step 58:  Execute Swipe  directions as edown and velocity emedium  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction edown and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiList , eDOWN , eMEDIUM  , ,80 ,obj_down)

' Step 59  Execute Swipe  directions as eup and velocity emedium and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity emedium and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiList  ,  eUP ,eMEDIUM ,  ,80,obj_up)

'' Step 60  Execute Swipe  directions as edown and velocity eslow  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eslow and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction eslow and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiList , eDOWN , eSLOW  , ,80 ,obj_down)


' Step 61  Execute Swipe  directions as eup and velocity eslow and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eslow  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity eslow and ending percentage 1-100"

strResult = VerifySwipe(objMobiList  ,eUP ,eSLOW ,  ,80,obj_up)





'' Step 64  Execute Swipe  directions as edown and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction eFast and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiList , eDOWN , eFAST  ,20 ,80 ,obj_down)


' Step 65  Execute Swipe  directions as eup and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity eFast and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiList  ,eUP ,eFAST , 20 ,80,obj_up)


'' Step 66  Execute Swipe  directions as edown and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction emedium and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiList , eDOWN , eMEDIUM  ,20 ,80 ,obj_down)


' Step 67  Execute Swipe  directions as eup and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity emedium and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiList  , eUP ,eMEDIUM , 20 ,80,obj_up)

'' Step 68  Execute Swipe  directions as edown and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi List with direction eslow and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiList , eDOWN , eSLOW  ,20 ,80 ,obj_down)


' Step 69  Execute Swipe  directions as eup and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiList." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi List with direction eup and velocity eslow and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiList  , eUP ,eSLOW , 20 ,80,obj_up)


'navigate to login screen
LogOut

' Step 6 Execute CheckProperty  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is not  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on mobilist." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should return False"
strResult = VerifyCheckProperty(objMobiList, "itemscount" ,strProperty , 15000, False)

' Step8:  Execute Exist  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is not  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on mobilist." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return false."
strResult =VerifyExist(objMobiList, False,15)

' Step14:  Execute WaitProperty  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute wait property when object is not visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on mobilist" & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value and should return False"
strProperty =objMobiList.GetTOProperty("itemscount")
strResult = VerifyWaitProperty(objMobiList, "itemscount", strProperty, 5000, False)


'******************************************************************************************************************************************************************

'Call function to end test iteration
EndTestIteration()





