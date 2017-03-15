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

' Step: Navigate to Controlrs screen
'Expected Result: Controls screen should be displayed

Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & ": " &_
"Navigate to Search screen" & VBNewLine
Environment("ExpectedResult") = "Controls screen should be displayed"


'Set object for Element
Set objMobiElement = MobiDevice("PhoneLookup").MobiElement("UILabel")

'Input  Parameters
arrTOProps = Array("visible" , "text" , "enabled")
arrTOValues = Array(True , "UILabel" , False)

arrROProps = Array("accessibilitylabel")
arrROvalues = Array("UILabel")

''Call function to createreport  template 
CreateReportTemplate

'Initial Setup
StrResult  = LoginAndNavigateToControlsPage("" , objMobiElement)


''*********************************************************************************************************************
' Step:Execute CaptureBitmap with .png file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the png file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiElement , "png")

' Step: Execute CaptureBitmap with .bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with  .bmp file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the bmp file to the defined location."
strResult =  VerifyCaptureBitmap(objMobiElement , "bmp")

' Step:Execute CaptureBitmap with .override.bmp file
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute CaptureBitmap with override .png file"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should throw an error messge for override messagefor .bmp file."
strResult =  VerifyCaptureBitmap(objMobiElement , "override_bmp")

' Step:Execute CaptureBitmap with .override.png file
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

strProperty = objMobiElement.GetTOProperty("visible")
strResult = VerifyCheckProperty(objMobiElement, "visible" ,strProperty , 5000, True)


' Step :   Execute Exist  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = VerifyExist(objMobiElement, True, 5)


' Step :  Execute GetTOProperties 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperties method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."

strResult = VerifyGetTOProperties(objMobiElement , arrTOProps)


' Step1 Execute GetROProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetROProperty method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
Set objMobiElement = MobiDevice("PhoneLookup").MobiElement("UILabel")
arrROProps = Array("accessibilitylabel")
arrROvalues = Array("UILabel")
strResult =VerifyGetROProperty(objMobiElement, arrROProps, arrROvalues)


' Step:  Execute GetTOProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTOProperty method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
strResult =VerifyGetTOProperty(objMobiElement, arrTOProps, arrTOValues)


' Step:  Execute ToString 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute ToString method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ToString on MobiElement." & VBNewLine
Environment("ExpectedResult") = "ToString should return the object type and class."
strResult = VerifyToString(objMobiElement)


' Step Execute WaitProperty  when object is visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Wait property method when object is visible "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiElement, "visible", "True", 5000, True)


' Step  Execute RefreshObject 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Refresh method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiElement." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = VerifyRefreshObject(objMobiElement)


' Step:  Execute GetVisibleText  with Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText with coordinates "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetvisibleText with coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetvisibleText  should return text within specified area."
strResult = VerifyGetVisibleText(objMobiElement , true)


' Step:  Execute GetVisibleText  without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetVisibleText without coordinates "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetvisibleText without coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetvisibleText  should return text within specified area."
strResult = VerifyGetVisibleText(objMobiElement ,False)


''Step : Execute Click with  Boundary coordinates
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Click method with boundary value"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "Click should work correctly."
'strResult = VerifyClick(objMobiElement, "withboundarycoords")
'
'
'MobiDevice("PhoneLookup").MobiButton("Controls").Click
'wait 3

''Step : Execute Click with  Zero coordinates
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Click method with Zero coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "Click should work correctly."
'strResult = VerifyClick(objMobiElement, "withzerovalues")
'
'MobiDevice("PhoneLookup").MobiButton("Controls").Click

'Step : Execute Click   Without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method Without coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withoutcoords")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 2

''Step : 'Execute Click with  Negative coordinates
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Click method with Negative coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Click on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "Click should work correctly."
'strResult = VerifyClick(objMobiElement, "withnegativecoords")


'Step : Execute Click with  x coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withxvalue")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 2

'Step : 'Execute Click with  y coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep

Environment("Description") = "Execute Click method with y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withyvalue")
MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 2

'Step:'Execute Click with  Valid X & Y coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with Valis x & y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = VerifyClick(objMobiElement, "withvalidvalue")
MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 2


'Step 22:  Execute Click with boundary coordinates at Top-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiElement for boundary co-ordinates at Top-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Left Corner."

blnStepRC = VerifyClick(objMobiElement, "withboundarycoordsTopLeft")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
'Step 23:  Execute Click with boundary coordinates at Top-Right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiElement for boundary co-ordinates at Top-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Top-Right Corner."

blnStepRC = VerifyClick(objMobiElement, "withboundarycoordsTopRight")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
'Step 24:  Execute Click with boundary coordinates at Bottom-Left corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiElement for boundary co-ordinates at Bottom-Left Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Left Corner."

blnStepRC = VerifyClick(objMobiElement, "withboundarycoordsBottomLeft")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
'Step 25:  Execute Click with boundary coordinates at bottom-right corner
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment.Value("Description") = "Execute Click on MobiElement for boundary co-ordinates at Bottom-Right Corner."
Environment("ExpectedResult") = "Click should work correctly with boundary co-ordinates at Bottom-Right Corner."

blnStepRC = VerifyClick(objMobiElement, "withboundarycoordsBottomRight")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3

'Step :Execute Click with  Random coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Click method with Random coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."

strResult = VerifyClick(objMobiElement, "withrandomcoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step :  Execute LongClick With Valid Lapse At Random Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse at random coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the MobiElement for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapserandomcoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3

' Step :  Execute LongClick With Valid Lapse At x Coordinates 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse with x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") ="LongClick should trigger press event on  the MobiElement  for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapsexcoords")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step:  Execute LongClick With Valid Lapse At yCoordinates 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse with y coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") ="LongClick should trigger press event on  the MobiElement for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapseycoords")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
'' Step:  Execute LongClick With Invalid Valid Lapse
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute LongClick With inValid Lapse"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "LongClick should throw an error"
'
'strResult = VerifyLongClick(objMobiElement ,"withinvalidlapsetime") 

 
' Step :  Execute LongClick with valid Lapse without coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick with VAlid Lapse without coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the MobiElement for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapsewithoutcoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step :  Execute LongClick With Valid Lapse At 0,0 Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse At 0,0 Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on  the MobiElement for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapsezerocoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3

' Step :  Execute LongClick With Valid Lapse At Boundary Coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse At Boundary Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "LongClick should trigger press event on MobiElement for the specified time"

strResult = VerifyLongClick(objMobiElement  , "validlapseboundarycoords")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step  Execute LongClick With Valid Lapse At y and xCoordinates 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute LongClick With Valid Lapse with y & x coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick with Boundary  coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") ="LongClick should trigger press event on  the MobiElement  for the specified time"

strResult = VerifyLongClick(objMobiElement  , "withvalidlapsevalidvalue")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step:  Execute DblClick  with Valid x & y  coordinates
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With valid X & Y coords"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withvalidvalues")


MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step:  Execute DblClick  withoutcoords  
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At without Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withoutcoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step :  Execute DblClick  withboundarycoords
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At Boundary Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withboundarycoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step:  Execute DblClick  withrandomcoords  
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At random Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withrandomcoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step :  Execute DblClick  withzercoords
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At zero Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withzercoords")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
'' Step:  Execute DblClick  withnegativecoords
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute DblClick With Valid Lapse At negative Coordinates"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "DblClick should throw an error"
'strResult = VerifyDblClick(objMobiElement  , "withnegativecoords")


' Step:  Execute DblClick   withonlyxcoord
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick At x Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withonlyxcoord")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
' Step :  Execute DblClick   withonlyycoord
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute DblClick With Valid Lapse At y Coordinates"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick with Boundary coordinates on MobiElement." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly"
strResult = VerifyDblClick(objMobiElement  , "withonlyycoord")

MobiDevice("PhoneLookup").MobiButton("Controls").Click
wait 3
'*********************************************************************************************************************

'*********************************************************************************************************************
' Step :  Execute GetTextlocation with text
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute GetTextlocation with text"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTextLocation  on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetTextLocation should return the co-ordinates"


MobiDevice("PhoneLookup").MobiList("lstControls").Scroll eTOP
wait 2
strResult = VerifyGetTextLocation(objMobiElement, "UILabel", True)

'*********************************************************************************************************************
'' Step :  Execute GetTextlocation without text
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute GetTextlocation without text"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute GetTextLocation  on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "GetTextLocation should  throw error message"
'
'strResult = VerifyGetTextLocation(objMobiElement, "", True)

'Open ScrollView screen
MobiDevice("PhoneLookup").MobiList("List").Select "UIScrollView"
Wait 3

Set objMobiElement1 = MobiDevice("PhoneLookup").MobiElement("eleScrollView")
objMobiElement1.highlight
'Step 7 : Execute ChildObjects with  non recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)with recursive value False"
blnResult = VerifyChildObjects(objMobiElement1, "nonrecursive" , 10 )

'Step 8 : Execute ChildObjects with recursive 
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =   " Verify child object count"
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any) with recursive value True"
blnResult = VerifyChildObjects(objMobiElement1, "recursive" ,10)



' Step   Execute Scroll  Bottom
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll bottom"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the MobiElement correctly towards down."

'Set objMobiElement = MobiDevice("PhoneLookup").MobiElement("ScrollView")
'MobiDevice("PhoneLookup").MobiList("List").Select "ScrollView"
'wait 3
Set objListControlBottom = MobiDevice("PhoneLookup").MobiElement("eleScrollView").MobiElement("eleBottom")


strResult = VerifyScroll(objMobiElement1, "bottom", objListControlBottom)


' Step :  Execute Scroll  TOP
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Scroll Top"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Scroll down on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Scroll should scroll the MobiElement correctly towards top."

Set objListControlTop = MobiDevice("PhoneLookup").MobiElement("eleScrollView").MobiElement("eleTop")
strResult = VerifyScroll(objMobiElement1, "top", objListControlTop)


'' Step :  Execute Scroll  without any  parameter
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("Description") = "Execute Scroll without any parameter"
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Scroll down on MobiElement." & VBNewLine
'Environment("ExpectedResult") = "Scroll should throw an error"
'
'strResult = VerifyScroll(objMobiElement1, "withoutparameter", "")
'

' Step :  Execute Swipe down
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe edown"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a swipe edown gesture on a Mobi Element"

'Swipe
Set obj_Bottom =MobiDevice("PhoneLookup").MobiElement("eleScrollView").MobiElement("eleBottom")

''Navigate to Scroll view screen
'MobiDevice("PhoneLookup").MobiList("List").Select "ScrollView"
'wait 3

strResult = VerifySwipe(objMobiElement1 ,eDOWN ,,,,obj_Bottom)

' Step :  Execute Swipe up
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe up"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a Swipe up gesture on a Mobi Element"

Set obj_Top = MobiDevice("PhoneLookup").MobiElement("eleScrollView").MobiElement("eleTop")
strResult = VerifySwipe(objMobiElement1 , eUP ,, 30, , objListControlBottom)


'' Step :  Execute Swipe  with directions as edown and velocity eFast
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as edow and velocity up"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity efast "
strResult = VerifySwipe(objMobiElement1 , eDOWN , eFAST  , , ,obj_Bottom)


' Step :  Execute  Swipe with directions as eup and velocity eFast
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity up"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity efast"

strResult = VerifySwipe(objMobiElement1  , eUP ,eFAST ,30  , ,obj_Top)



'' Step :  Execute Swipe  with directions as edown and velocity emedium
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as edow and velocity emedium"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity emedium "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eMEDIUM  , ,  ,obj_Bottom)


' Step :  Execute  Swipe with directions as eup and velocity emedium
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity emedium"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium"

strResult = VerifySwipe(objMobiElement1  ,eUP ,eMEDIUM , 30 , ,obj_Top)


'' Step :  Execute Swipe  with directions as edown and velocity eslow
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execut swipe with direction as edow and velocity emedium"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity eslow "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eSLOW , ,  ,obj_Bottom)


' Step :  Execute  Swipe with directions as eup and velocity eslow
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut swipe with direction as eup and velocity eslow"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow"

strResult = VerifySwipe(objMobiElement1  ,eUP ,eSLOW ,30  , ,obj_Top)

'' Step :  Execute Swipe  directions as edown and velocity eFast  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eFAST  ,20 ,  ,obj_Bottom)


' Step   Execute Swipe  directions as eup and velocity eFast and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eFast  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity efast and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement1  ,  eUP ,eFAST , 20 , ,obj_Top)

'' Step :  Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity emedium and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eMEDIUM  ,20 ,  ,obj_Bottom)


' Step   Execute Swipe  directions as eup and velocity emedium and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity emedium  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement1  ,  eUP ,eMEDIUM , 20 ,,obj_Top)

'' Step   Execute Swipe  directions as edown and velocity eslow  and starting percentage as 0-99
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eslow and velocity eFast  and starting percentage as 0-99"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eslow and velocity efast and starting percentage 0-99 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eSLOW ,20 , ,obj_Bottom)


' Step   Execute Swipe  directions as eup and velocity eslow and starting percentage as 0-99
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eslow  and starting percentage as 0-99"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow and starting percentage 0-99"

strResult = VerifySwipe(objMobiElement1  ,  eUP ,eSLOW , 20 , ,obj_Top)


'' Step : Execute Swipe  without any parameter
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute Swipe  without any parameter"
'Environment("ExpectedResult") = "S wipe should throw an error"
'strResult = VerifySwipe(objMobiElement1, , , , ,objMobiElement1 )
'
'' Step : Execute Swipe  with valid direction and invalid  velocity
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "ExecuteSwipe  with valid direction and invalid  velocity"
'Environment("ExpectedResult") = "Swipe should throw an error"
'
'strResult = VerifySwipe(objMobiElement1  , eDOWN ,afsa ,, ,obj_Bottom)
'
'' Step :  Execute Swipe  directions as edown and velocity eFast  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eFAST  , , 80 ,obj_Bottom)

' Step   Execute Swipe  directions as eup and velocity eFast and and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eFast  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity efast and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1  ,  eUP ,eFAST , 30 , 80,obj_Top)

'' Step :  Execute Swipe  directions as edown and velocity emedium  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction edown and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1 , eDOWN , eMEDIUM  , ,80 ,obj_Bottom)

' Step  Execute Swipe  directions as eup and velocity emedium and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity emedium and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1  ,  eUP ,eMEDIUM ,30  ,80,obj_Top)

'' Step   Execute Swipe  directions as edown and velocity eslow  and ending percentage 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as eslow and velocity eFast  and ending percentage 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eslow and velocity efast and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eSLOW  , ,80 ,obj_Bottom)


' Step  Execute Swipe  directions as eup and velocity eslow and ending percentage 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execut Swipe  directions as edown and velocity eslow  and ending percentage 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1  ,eUP ,eSLOW , 30 ,80,obj_Top)


'' Step  Execute Swipe  with valid direction,velocity and invalid start percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute Swipe  with valid direction,velocity and invalid start percentage"
'Environment("ExpectedResult") = "Swipe should throw an error"
'
'strResult = VerifySwipe(objMobiElement1  ,  eUP ,eSLOW , 10.57 ,,objMobiElement1)
'
'' Step   Execute Swipe  with valid direction,velocity and invalid end percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") = "Execute Swipe  with valid direction,velocity and invalid end percentage"
'Environment("ExpectedResult") = "Swipe should throw an error"
'
'strResult = VerifySwipe(objMobiElement1  ,  eUP ,eSLOW ,  ,10.57,objMobiElement1)
'

'' Step  Execute Swipe  directions as edown and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eFast and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eFAST  ,20 ,80 ,obj_Bottom)


' Step   Execute Swipe  directions as eup and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity eFast and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eFast and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1  ,eUP ,eFAST , 20 ,80,obj_Top)


'' Step   Execute Swipe  directions as edown and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction emedium and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eMEDIUM  ,20 ,80 ,obj_Bottom)


' Step   Execute Swipe  directions as eup and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity emedium and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity emedium and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1  , eUP ,eMEDIUM , 20 ,80,obj_Top)

'' Step   Execute Swipe  directions as edown and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
''#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Swipe  directions as edown and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Simulates a gesture on  Mobi Element with direction eslow and velocity efast  and  starting percentage as 0-99 and ending percentage 1-100 "

strResult = VerifySwipe(objMobiElement1 , eDOWN , eSLOW  ,20 ,80 ,obj_Bottom)


' Step   Execute Swipe  directions as eup and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Swipe on MobiElement." & VBNewLine
Environment("Description") = "Execute Swipe  directions as eup and velocity eslow and  starting percentage as 0-99 and ending percentage as 1-100"
Environment("ExpectedResult") = "Simulates a gesture on a Mobi Element with direction eup and velocity eslow and  starting percentage as 0-99  and ending percentage 1-100"

strResult = VerifySwipe(objMobiElement1  , eUP ,eSLOW , 20 ,80,obj_Top)


'' Step   Execute Swipe  with valid direction ,velocity  and negative start percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") ="Execute Swipe  with valid direction ,velocity  and negative start percentage"
'Environment("ExpectedResult") = "Swipe  should throw an error"
'
'strResult = VerifySwipe(objMobiElement1  ,eUP ,eSLOW , -20 ,,objMobiElement1)
'
'' Step   Execute Swipe  with valid direction ,velocity  and negative end  percentage
''#######################################################
'intStep = intStep+1
'Environment("StepName") = "Step" & intStep
'Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
'"Execute Swipe on MobiElement." & VBNewLine
'Environment("Description") ="Execute Swipe  with valid direction ,velocity  and negative end percentage"
'Environment("ExpectedResult") = "Swipe  should throw an error"
'
'strResult = VerifySwipe(objMobiElement1  , eUP ,eSLOW ,  ,-80,objMobiElement1)
'
'wait 2
'While Not  (FetchHeaderElement = "Controls")
'			'Navigate back to  Object screen
'		MobiDevice("PhoneLookup").MobiButton("index:=0").Click
'		Wait 1
'Wend
'

' Step  :Execute SetTOProperty
'##########################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute SetToProperty"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."

strResult = VerifySetTOProperty(objMobiElement, arrTOProps)

' Step :  Execute GetScrolledText 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") =  "Execute GetScrolledText method"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetScrolledText on MobiElement." & VBNewLine
Environment("ExpectedResult") = "GetScrolledText should return text of the scrolled window"

''Navigate to Scroll view screen
'MobiDevice("PhoneLookup").MobiList("List").Select "UIScrollView"
'wait (2)

strResult = VerifyGetScrollText(objMobiElement1 ,"element" , False , wheelcount  , ""  )

'LogOut
LogOut

' Step  Execute CheckProperty  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute check property when object is not  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should return false and  wait for the property to attain a value and report the result."
strResult = VerifyCheckProperty(objMobiElement, "visible" ,strProperty , 5000, False)


' Step8:  Execute Exist  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute Exist method when object is not  visible"
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiElement." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly and return false."
strResult = VerifyExist(objMobiElement, False, 5)

' Step14:  Execute WaitProperty  when object is not visible
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("Description") = "Execute wait property method when object is not visible "
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiElement." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = VerifyWaitProperty(objMobiElement, "visible", "True", 5000, False)


'******************************************************************************************************************************************************************

'Call function to end test iteration
EndTestIteration()




