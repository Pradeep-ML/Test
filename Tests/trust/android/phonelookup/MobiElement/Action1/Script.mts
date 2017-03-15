''Verify MobiElement  object methods

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

'Set object for Element
Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("AbsoluteLayout")

''Call function to createreporttemplare
CreateReportTemplate()

'Initial Setup
StrResult = NavigateScreenOnPhoneLookup("Controls" , objMobiElement , "")


'*********************************************************************************************************************
' Step1:Execute CaptureBitmap with .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the png file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiElement , "png")


' Step2:Execute CaptureBitmap with .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with  .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the bmp file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiElement , "bmp")


' Step3:Execute CaptureBitmap with .override.bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with override .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw an error messge for override messagefor .bmp file."
strResult =  VerifyCaptureBitmap(objMobiElement , "override_bmp")

' Step4:Execute CaptureBitmap with .override.png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") =  "CaptureBitmap should throw an error messge for override message for .png  file."
strResult =  VerifyCaptureBitmap(objMobiElement , "override_png")

' Step 5 Execute CheckProperty  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strResult = VerifyCheckProperty(objMobiElement, "visible" ,True , 5000, True)


' Step 119 :Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
arrProps = Array("enabled","text")
strResult = VerifySetTOProperty(MobiElement, arrProps)

' Step7:  Execute Exist  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = VerifyExist(objMobiElement, True, 5)


' Step9:  Execute GetTOProperties 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("visible","text","id")
strResult = VerifyGetTOProperties(objMobiElement, arrProps)


' Step10  Execute GetROProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
arrROProps = Array("enabled","text","id")
arrROvalue= Array (True,"AbsoluteLayout","-1")
strResult =VerifyGetROProperty(objMobiElement, arrROProps, arrROvalue)


' Step11:  Execute GetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
arrProps = Array("enabled","text","id")
arrvalue= Array (True,"AbsoluteLayout","-1")
strResult =VerifyGetTOProperty(objMobiElement, arrProps, arrvalue)


' Step12:  Execute ToString 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ToString on MobiElement." & VBNewLine
Environment("ExpectedResult") = "ToString should return the object type and class."
strResult = VerifyToString(objMobiElement)

' Step13  Execute WaitProperty  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Wait property when object is visible method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiElement, "visible", True, 5000, True)


' Step16  Execute RefreshObject 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Refresh method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiElement." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiElement)

' Step17:  Execute GetVisibleText  with Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText with coordinates "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetvisibleText with coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetvisibleText  should return text within specified area."
strResult = VerifyGetVisibleText(objMobiElement , true)

' Step18:  Execute GetVisibleText  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText without coordinates "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetvisibleText without coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetvisibleText  should return text within specified area."
strResult = VerifyGetVisibleText(objMobiElement ,False)


' 'Step 13:  Execute Click  with boundary coordinates at Top-Left corner
''#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Click : Execute Click with boundary coordinates at Top-Left corner on MobiList."
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiEdit." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
blnFlag = VerifyClick(objMobiElement ,"withboundarycoordsTopLeft")
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
blnFlag = VerifyClick(objMobiElement,"withboundarycoordsTopRight")
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
blnFlag = VerifyClick(objMobiElement,"withboundarycoordsBottomLeft")
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
blnFlag = VerifyClick(objMobiElement,"withboundarycoordsBottomRight")
GoToScreen "Controls"




'22'Execute Click   Without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method Without coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withoutcoords")

GoToScreen "Controls"





'24'Execute Click with  x coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withxvalue")

GoToScreen "Controls"

'25'Execute Click with  y coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withyvalue")

GoToScreen "Controls"


'26'Execute Click with  Valid X & Y coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with Valis x & y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withvalidvalue")

GoToScreen "Controls"


'27'Execute Click with  Random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with Random coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."

strResult = VerifyClick(objMobiElement, "withrandomcoords")

GoToScreen "Controls"



 
' Step 28:  Execute LongClick with VAlid Lapse without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick with VAlid Lapse without coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapsewithoutcoords")
wait 2

GoToScreen "Controls"

' Step 29:  Execute LongClick With Valid Lapse At 0,0 Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse At 0,0 Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapsezerocoords")
wait 2

GoToScreen "Controls"

' Step 30:  Execute LongClick With Valid Lapse At Boundary Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse At Boundary Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapseboundarycoords")
wait 2

GoToScreen "Controls"


' Step 31: Execute LongClick With Valid Lapse At y and xCoordinates 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse with y & x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") ="LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "withvalidlapsevalidvalue")

GoToScreen "Controls"


' Step32:  Execute DblClick  with Valid x & y  coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With valid X & Y coords"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withvalidvalues")

GoToScreen "Controls"

' Step33  Execute DblClick  withoutcoords  
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At without Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withoutcoords")

GoToScreen "Controls"

' Step34:  Execute DblClick  withboundarycoords
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At Boundary Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withboundarycoords")

GoToScreen "Controls"

' Step35:  Execute DblClick  withrandomcoords  
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At random Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withrandomcoords")

GoToScreen "Controls"

' Step36:  Execute DblClick  withzercoords
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At zero Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withrandomcoords")

GoToScreen "Controls"




' Step38:  Execute DblClick   withonlyxcoord
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick At x Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withonlyxcoord")

GoToScreen "Controls"

' Step39:  Execute DblClick   withonlyycoord
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At y Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withonlyycoord")

GoToScreen "Controls"

'*********************************************************************************************************************
' Step40:  Execute GetScrolledText 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description")  = "Execute GetScrolledText method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetScrolledText on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetScrolledText should return text of the scrolled window"

Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("ScrollView")

'Navigate to Scroll view screen
MobiDevice("Phone Lookup").MobiList("List").Select "ScrollView"
wait (2)

strResult = VerifyGetScrollText(objMobiElement,"mobielement", "false" , "" , "")
wait 2

'Returned back to controls screen
MobiDevice("Phone Lookup").ButtonPress eBACK
wait 3

'*********************************************************************************************************************
'Step 7:  Execute  ChildObjects recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute childobject method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on mobilist." & VBNewLine

Environment("ExpectedResult") = "Return child object recursively in the application"
'blnFlag = VerifyChildObjects(objMobiElement  ,"recursive",26)
blnFlag = VerifyChildObjects(objMobiElement, "nonrecursive" , 0)

 'Step 7:  Execute  ChildObjects non recursively
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute childobject method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on mobilist." & VBNewLine

Environment("ExpectedResult") = "Return child object non recursively in the application"
'blnFlag = VerifyChildObjects(objMobiElement,"recursive",100)
blnFlag = VerifyChildObjects(objMobiElement, "nonrecursive" , 100)


' Step41:  Execute GetTextlocation with text
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTextlocation with text"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetScrolledText on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetTextLocation should return the co-ordinates"

Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("AbsoluteLayout")
MobiDevice("Phone Lookup").MobiList("List").Select "AbsoluteLayout"
Wait 1
MobiDevice("Phone Lookup").ButtonPress eBACK
Wait 2

strText  =  objMobiElement.GetVisibleText
strResult = VerifyGetTextLocation(objMobiElement, strText , True)

'*********************************************************************************************************************



' Step 43 Execute Scroll  Bottom
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll bottom"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards down."

Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("ScrollView")
MobiDevice("Phone Lookup").MobiList("List").Select "ScrollView"
wait 3
Set objListControlBottom = MobiDevice("Phone Lookup").MobiElement("eleBottom")

strResult = VerifyScroll(objMobiElement, "bottom", objListControlBottom)


' Step 44:  Execute Scroll  TOP
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll Top"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards top."

Set objListControlTop = MobiDevice("Phone Lookup").MobiElement("eleTop")
strResult = VerifyScroll(objMobiElement, "top", objListControlTop)

wait 2
'Returned back to controls screen
MobiDevice("Phone Lookup").ButtonPress eBACK
wait 3


' Step 45:  Execute Scroll  Right
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll Right"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards Right."

'Navigate to Horizontal scroll view screen
 MobiDevice("Phone Lookup").MobiList("List").Select "HorizontalScrollView"
wait 2
Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("HorizontalScrollView")

Set objListControlRight = MobiDevice("Phone Lookup").MobiElement("eleTop")
strResult = VerifyScroll(objMobiElement, "right", objListControlRight)


' Step 46:  Execute Scroll  left
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll Left"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the list correctly towards left."

Set objListControlleft = MobiDevice("Phone Lookup").MobiElement("eleTop")
strResult = VerifyScroll(objMobiElement, "left", objListControlleft)




' Step 48:  Execute Swipe down
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe edown"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a swipe edown gesture on a Mobi Element"

'Swipe
Set objBottom = MobiDevice("Phone Lookup").MobiElement("eleBottom_Swipe")

Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("ScrollView")
MobiDevice("Phone Lookup").ButtonPress eBACK
Wait 3
'Navigate to Scroll view screen
MobiDevice("Phone Lookup").MobiList("List").Select "ScrollView"
wait 3

strResult = VerifySwipe(objMobiElement ,eDOWN ,,,,objBottom)

' Step 49:  Execute Swipe up
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe up"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a Swipe up gesture on a Mobi Element"

Set objTop = MobiDevice("Phone Lookup").MobiElement("elescrollviewtop")
strResult = VerifySwipe(objMobiElement , eUP , , , , objTop)


'' Step 50:  Execute Swipe  with directions as edown and velocity eFast
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as edow and velocity up"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity efast "

Set objBottom = MobiDevice("Phone Lookup").MobiElement("elescrollviewdown")

strResult = VerifySwipe(objMobiElement , eDOWN , eFAST  , , ,objBottom)


' Step 51:  Execute  Swipe with directions as eup and velocity eFast
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity up"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity efast"



strResult = VerifySwipe(objMobiElement  , eUP ,eFAST ,  , ,objTop)



'' Step 52:  Execute Swipe  with directions as edown and velocity emedium
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as edow and velocity emedium"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity emedium "

strResult = VerifySwipe(objMobiElement , eDOWN , eMEDIUM  , ,  ,objBottom)


' Step 53:  Execute  Swipe with directions as eup and velocity emedium
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity emedium"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium"

strResult = VerifySwipe(objMobiElement  ,eUP ,eMEDIUM ,  , ,objTop)


'' Step 54:  Execute Swipe  with directions as edown and velocity eslow
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as edow and velocity emedium"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity eslow "

strResult = VerifySwipe(objMobiElement , eDOWN , eSLOW , ,  ,objBottom)


' Step 55:  Execute  Swipe with directions as eup and velocity eslow
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity eslow"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow"

strResult = VerifySwipe(objMobiElement  ,eUP ,eSLOW ,  , ,objTop)

'' Step 56:  Execute Swipe  directions as edown and velocity eFast  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement , eDOWN , eFAST  ,20 ,  ,objBottom)


' Step 57  Execute Swipe  directions as eup and velocity eFast and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eFast  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity efast and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement  ,  eUP ,eFAST , 20 , ,objTop)

'' Step 58:  Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity emedium and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement , eDOWN , eMEDIUM  ,20 ,  ,objBottom)


' Step 59  Execute Swipe  directions as eup and velocity emedium and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement  ,  eUP ,eMEDIUM , 20 ,,objTop)

'' Step 60  Execute Swipe  directions as edown and velocity eslow  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eslow and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eslow and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement , eDOWN , eSLOW ,20 , ,objBottom)


' Step 61  Execute Swipe  directions as eup and velocity eslow and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eslow  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement  ,  eUP ,eSLOW , 20 , ,objTop)




'' Step 64:  Execute Swipe  directions as edown and velocity eFast  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eDOWN , eFAST  , , 80 ,objBottom)

' Step 65  Execute Swipe  directions as eup and velocity eFast and and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eFast  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity efast and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,  eUP ,eFAST ,  , 80,objTop)

'' Step 66:  Execute Swipe  directions as edown and velocity emedium  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement , eDOWN , eMEDIUM  , ,80 ,objBottom)

' Step 67  Execute Swipe  directions as eup and velocity emedium and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity emedium and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,  eUP ,eMEDIUM ,  ,80,objTop)

'' Step 68  Execute Swipe  directions as edown and velocity eslow  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eslow and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eslow and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eDOWN , eSLOW  , ,80 ,objBottom)


' Step 69  Execute Swipe  directions as eup and velocity eslow and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eslow  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,eUP ,eSLOW ,  ,80,objTop)





'' Step 72  Execute Swipe  directions as edown and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eFast and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eDOWN , eFAST  ,20 ,80 ,objBottom)


' Step 73  Execute Swipe  directions as eup and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eFast and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,eUP ,eFAST , 20 ,80,objTop)


'' Step 74  Execute Swipe  directions as edown and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction emedium and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eDOWN , eMEDIUM  ,20 ,80 ,objBottom)


' Step 75  Execute Swipe  directions as eup and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  , eUP ,eMEDIUM , 20 ,80,objTop)

'' Step 76  Execute Swipe  directions as edown and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eslow and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eDOWN , eSLOW  ,20 ,80 ,objBottom)


' Step 77  Execute Swipe  directions as eup and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  , eUP ,eSLOW , 20 ,80,objTop)




'Back To Control Screen and navigate to horizontalScrollview
wait 2
MobiDevice("Phone Lookup").ButtonPress eBACK
wait 4
MobiDevice("Phone Lookup").MobiList("List").Select "horizontalscrollview"
wait 1


' Step 80:  Execute Swipe eright
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eright"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a swipe eright gesture on a Mobi Element"

'Swipe
Set obj_Right = MobiDevice("Phone Lookup").MobiElement("ScrollView_Right")
Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("HorizontalScrollView")

strResult = VerifySwipe(objMobiElement  , eRIGHT,  , , ,obj_Right)

' Step 81:  Execute Swipe eleft
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe eleft"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a Swipe eleft  gesture on a Mobi Element"

Set objListControlleft = MobiDevice("Phone Lookup").MobiElement("ScrollView_Left")

strResult = VerifySwipe(objMobiElement , eLEFT  , , , , objListControlleft)


'' Step 82:  Execute Swipe  with directions as eright and velocity eFast
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as eright and velocity up"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity efast "

'Set obj_down = MobiDevice("Phone Lookup").MobiElement("eleBottom_Swipe")

strResult = VerifySwipe(objMobiElement , eRIGHT , eFAST , , ,obj_Right)


' Step 83:  Execute  Swipe with directions as eleft and velocity eFast
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eleft and velocity up"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity efast"

Set obj_up = MobiDevice("Phone Lookup").MobiElement("eleTop")

strResult = VerifySwipe(objMobiElement  ,  eLEFT ,eFAST ,  ,,objListControlleft)



'' Step 84:  Execute Swipe  with directions as eright and velocity emedium
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as eright and velocity emedium"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity emedium "

strResult = VerifySwipe(objMobiElement , eRIGHT , eMEDIUM , , ,obj_Right)


' Step 85:  Execute  Swipe with directions as eleft and velocity emedium
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eleft and velocity emedium"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity emedium"

strResult = VerifySwipe(objMobiElement  ,eLEFT ,eMEDIUM ,  , ,objListControlleft)

	
'' Step 86:  Execute Swipe  with directions as eright and velocity eslow
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as eright and velocity emedium"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity eslow "

strResult = VerifySwipe(objMobiElement , eRIGHT , eSLOW  , , ,obj_Right)


' Step 87:  Execute  Swipe with directions as eleft and velocity eslow
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eleft and velocity eslow"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity eslow"

strResult = VerifySwipe(objMobiElement  ,eLEFT ,eSLOW , ,,objListControlleft)

'' Step 88:  Execute Swipe  directions as eright and velocity eFast  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement , eRIGHT , eFAST  ,20 , ,obj_Right)


' Step 89  Execute Swipe  directions as eleft and velocity eFast and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as eleft and velocity eFast  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity efast and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement  ,eLEFT ,eFAST , 20 , ,objListControlleft)

'' Step 90:  Execute Swipe  directions as eright and velocity emedium  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity emedium  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity emedium and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement , eRIGHT , eMEDIUM ,20 , ,obj_Right)


' Step 91  Execute Swipe  directions as eleft and velocity emedium and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as eleft and velocity emedium  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity emedium and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement  ,  eLEFT ,eMEDIUM , 20 , ,objListControlleft)

'' Step 92  Execute Swipe  directions as eright and velocity eslow  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement , eRIGHT , eSLOW ,20 , ,obj_Right)


' Step 93  Execute Swipe  directions as eLEFT and velocity eslow and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as eLEFT and velocity eslow  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eLEFT and velocity eslow and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement  ,  eLEFT ,eSLOW , 20 ,,objListControlleft)

'' Step 94:  Execute Swipe  directions as eright and velocity eFast  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eRIGHT, eFAST  , , 80 ,obj_Right)

' Step 95  Execute Swipe  directions as eleft and velocity eFast and and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as eleft and velocity eFast  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity efast and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,  eLEFT ,eFAST ,  , 80, objListControlleft)

'' Step 96:  Execute Swipe  directions as eright and velocity emedium  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity emedium  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement , eRIGHT ,eMEDIUM   , , 80 ,obj_Right)

' Step 97  Execute Swipe  directions as eleft and velocity emedium and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as eleft and velocity emedium and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  , eLEFT ,eMEDIUM ,  ,80,objListControlleft)

'' Step 98  Execute Swipe  directions as eright and velocity eslow  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eRIGHT , eSLOW  ,"" ,80 ,obj_Right)


' Step 99  Execute Swipe  directions as eleft and velocity eslow and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as eleft and velocity eslow  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity eslow and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,eLEFT ,eSLOW ,  ,80,objListControlleft)



'' Step 100  Execute Swipe  directions as eright and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement ,eRIGHT , eFAST  ,20 ,80 ,obj_Right)


' Step 101  Execute Swipe  directions as eleft and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eleft and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity eFast and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,eLEFT ,eFAST , 20 ,80,objListControlleft)


'' Step 102  Execute Swipe  directions as eright and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity emedium  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eRIGHT , eMEDIUM ,20 ,80 ,obj_Right)


' Step 103  Execute Swipe  directions as eleft and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eleft and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity emedium and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  , eLEFT ,eMEDIUM , 20 ,80,objListControlleft)

'' Step 104  Execute Swipe  directions as eright and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eright and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eright and velocity eslow  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement , eRIGHT , eSLOW  ,20 ,80 ,obj_Right)


' Step 105  Execute Swipe  directions as eleft and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eleft and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eleft and velocity eslow and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement  ,  eLEFT ,eSLOW , 20 ,80,objListControlleft)

'returning back to controls screen
wait 2
MobiDevice("Phone Lookup").ButtonPress eBACK
wait 4

'' Step 106:Execute ClickColor  with color as a # value
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with color as a #value"
'Environment("ExpectedResult") = "Click color should click on the color passed"
'
'MobiDevice("Phone Lookup").MobiList("List").Select "AbsoluteLayout"
'Wait 2
'MobiDevice("Phone Lookup").ButtonPress eBACK
'Wait 2
'
'Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("AbsoluteLayout")
'
'strResult = VerifyClickColor(objMobiElement, "validwithoutdiff", "#EAFFFF", "")
'
'GoToScreen "Controls"
'
'' Step 107:Execute ClickColor  with color as a string value
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with color as a string value"
'Environment("ExpectedResult") = "Click color should click on the color passed"
'
'strResult = VerifyClickColor(objMobiElement, "validwithoutdiff", "White", "")
'
'GoToScreen "Controls"
'
'' Step 108:Execute ClickColor  with invalidcolor as a # value
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with invalidcolor as a #value"
'Environment("ExpectedResult") = "Click color should not click on the color passed"
'
'
'Set objMobiElement = MobiDevice("Phone Lookup").MobiElement("AbsoluteLayout")
'
'strResult = VerifyClickColor(objMobiElement," invalidcolor", "#fdsgsdfhgghfdgergdfxgtehgdfxgtd", "")
'
'
'' Step 109:Execute ClickColor  with invalidcolor as a string value
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with invalidcolor as a string value"
'Environment("ExpectedResult") = "Click color should not click on the color passed"
'
'strResult = VerifyClickColor(objMobiElement, "invalidcolor", "grratttg", "")
'
'
'' Step 110:Execute ClickColor  with color as a # value and valid allowable difference
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with color as a #value and valid allowable difference"
'Environment("ExpectedResult") = "Click color should click on the color passed"
'
'strResult = VerifyClickColor(objMobiElement, "validwithdiff", "#EAFFFF", 10)
'
'GoToScreen "Controls"
'
'' Step 111:Execute ClickColor  with color as a string value and valid allowable difference
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with color as a string value and valid allowable difference"
'Environment("ExpectedResult") = "Click color should click on the color passed"
'
'strResult = VerifyClickColor(objMobiElement, "validwithdiff", "White", 20)
'
'GoToScreen "Controls"
'
'' Step 112:Execute ClickColor  with Color as #value & In-Valid Allowable Difference
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute ClickColor  with Color as String & In-Valid Allowable Difference"
'Environment("ExpectedResult") = "Click color should not click on the color passed"
'
'strResult = VerifyClickColor(objMobiElement," invaliddiff", "#EAFFFF", -20)
'
'
'' Step 113:Execute ClickColor  with color as a string value and invalidallowablediff
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor with color as a string value and invalid allowable diff"
'Environment("ExpectedResult") = "Click color should not click on the color passed"
'
'strResult = VerifyClickColor(objMobiElement, "invaliddiff", "White", -20*10)
'
'
'' Step 114:Execute ClickColor  without any parameter
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute clickcolor without any parameter"
'Environment("ExpectedResult") = "Click color should throw an error"
'
'strResult = VerifyClickColor(objMobiElement, withoutparameter, "", "")


' Step 115:  Execute LongClick With Valid Lapse At Random Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse at random coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapserandomcoords")

GoToScreen "Controls"




' Step 117:  Execute LongClick With Valid Lapse At x Coordinates 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse with x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") ="LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapsexcoords")

GoToScreen "Controls"


' Step 118:  Execute LongClick With Valid Lapse At yCoordinates 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse with y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") ="LongClick should trigger press event on  the mobile device window for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapseycoords")


' Step 6 Execute CheckProperty  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is not  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should return false and  wait for the property to attain a value and report the result."
strResult = VerifyCheckProperty(objMobiElement, "visible" ,True , 15000, False)

' Step8:  Execute Exist  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is not  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return false."
strResult =VerifyExist(objMobiElement, False, 15)

' Step14:  Execute WaitProperty  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute wait property when object is not visible method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiElement, "visible", True, 15000, False)

'******************************************************************************************************************************************************************

'Call function to end test iteration
EndTestIteration()












