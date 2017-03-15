'##########################################################################################################
' Objective: Login to TapToZoom application and  execute all mobiimage methods.
' Test Description: Execute all methods for MobiImage 
' Steps:
' Step1: Execute CaptureBitmap
' Step2: Execute CheckProperty
' Step3: Execute ChildObjects
' Step4: Execute Click  without coordinates
' Step5: Execute Click  with coordinates
' Step6: Execute DblClick  without coordinates
' Step7: Execute DblClick  with coordinates
' Step8: Execute Exist
' Step9: Execute GetROProperty
' Step10: Execute GetTOProperties
' Step11: Execute GetTOProperty
' Step12: Execute GetVisibleText  without coordinates
' Step13: Execute GetVisibleText  with coordinates
' Step14: Execute GetTextLocation
' Step15: Execute LongClick  without coordinates
' Step16: Execute LongClick  with coordinates
' Step17: Execute RefreshObject
' Step18: Execute SetToProperty
' Step19: Execute ToString
' Step20: Execute WaitProperty
' ##########################################################################################################

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
Environment("StepName") = ""
'#######################################################

'#######################################################
'Set object for mobiimage

Set objMobiImage = MobiDevice("TapToZoom").MobiElement("Element").MobiImage("Image")


'#######################################################

' Step1:  Execute CaptureBitmap 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CaptureBitmap on MobiImage." & VBNewLine
Environment("ExpectedResult") = "CaptureBitmap should capture the screenshot and save the file to the defined location."
strResult = CStr(VerifyCaptureBitmap(objMobiImage))


' Step2:  Execute CheckProperty 
'#######################################################
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute CheckProperty on MobiImage." & VBNewLine
Environment("ExpectedResult") = "CheckProperty should wait for the property to attain a value and report the result."
strResult = strResult & CStr(VerifyCheckProperty(objMobiImage, "id", "100", 5000, True))


' Step3:  Execute  ChildObjects
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute ChildObjects on MobiImage." & VBNewLine
Environment("ExpectedResult") = "ChildObjects should return the count of children (if any)."
strResult = strResult & CStr(VerifyChildObjects(objMobiImage))



' Step4:  Execute Click  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiImage." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = strResult & CStr(VerifyClick(objMobiImage, False))

' Step5:  Execute Click  with coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Click on MobiImage." & VBNewLine
Environment("ExpectedResult") = "Click should work correctly."
strResult = strResult & CStr(VerifyClick(objMobiImage, True))



' Step6:  Execute DblClick  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick on MobiImage." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
strResult = strResult & CStr(VerifyDblClick(objMobiImage, False))


' Step7:  Execute DblClick  with coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute DblClick on MobiImage." & VBNewLine
Environment("ExpectedResult") = "DblClick should work correctly."
strResult = strResult & CStr(VerifyDblClick(objMobiImage, True))

' Step8:  Execute Exist 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute Exist on MobiImage." & VBNewLine
Environment("ExpectedResult") = "Exist should work correctly."
strResult = strResult & CStr(VerifyExist(objMobiImage, True, 5))


' Step9:  Execute GetROProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetROProperty on MobiImage." & VBNewLine
Environment("ExpectedResult") = "GetROProperty should get the correct run-time value for a property."
'strResult = strResult & CStr(VerifyGetROProperty(objMobiImage))

' Step10:  Execute GetTOProperties 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperties on MobiImage." & VBNewLine
Environment("ExpectedResult") = "GetTOProperties should get all properties for an object that are used for description."
arrProps = Array("value" , "id")
strResult = strResult & CStr(VerifyGetTOProperties(objMobiImage, arrProps, strMissingProperties))

' Step11:  Execute GetTOProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTOProperty on MobiImage." & VBNewLine
Environment("ExpectedResult") = "GetTOProperty should return the correct value for the test object property."
strResult = strResult & CStr(VerifyGetTOProperty(objMobiImage, arrProps, strNotFound))

' Step12:  Execute GetVisibleText  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MobiImage." & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
strResult = strResult & CStr(VerifyGetVisibleText(objMobiImage,False))

 'Step13:  Execute GetVisibleText  with coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetVisibleText on MObiIMage " & VBNewLine
Environment("ExpectedResult") = "GetVisibleText should return correct text after OCRing."
strResult = strResult & CStr(VerifyGetVisibleText(objMobiImage,True))


 'Step14:  Execute GetTextLocation
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute GetTextLocation on MobiImage." & VBNewLine
Environment("ExpectedResult") = "GetTextLocation should return the correct value for the test object property."
strText = objMobiImage.GetVisibleText()
strResult = strResult & CStr(VerifyGetTextLocation(objMobiImage , strText  , True))


' Step15:  Execute LongClick  without coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick on MobiImage." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
strResult = strResult & CStr(VerifyLongClick(objMobiImage , False))

' Step16:  Execute LongClick  with coordinates
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute LongClick on MobiImage." & VBNewLine
Environment("ExpectedResult") = "LongClick should work correctly."
strResult = strResult & CStr(VerifyLongClick(objMobiImage , True))

' Step17:  Execute RefreshObject
'#######################################################
'
intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute RefreshObject  on MobiImage." & VBNewLine
Environment("ExpectedResult") = "RefreshObject re-identifies the object in the application"
strResult = strResult & CStr(VerifyRefreshObject(objMobiImage))


' Step18:  Execute SetTOProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute SetTOProperty on MobiImage." & VBNewLine
Environment("ExpectedResult") = "SetTOProperty should set the value of the test object property."
strResult = strResult & CStr(VerifySetTOProperty(objMobiImage, arrProps, strSetFailedFor))

' Step19:  Execute TOString 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute TOString on MobiImage." & VBNewLine
Environment("ExpectedResult") = "TOString should return the object type and class."
strResult = strResult & CStr(VerifyTOString(objMobiImage))

' Step20:  Execute WaitProperty 
'#######################################################

intStep = intStep+1
Environment("StepName") = "Step" & intStep
Environment("StepsToReproduce") = Environment("StepsToReproduce") & Environment("StepName") & "." & intStep & ": " &_
"Execute WaitProperty on MobiImage." & VBNewLine
Environment("ExpectedResult") = "WaitProperty should wait for the property to attain a value but shouldn't report the result."
strResult = strResult & CStr(VerifyWaitProperty(objMobiImage, "id", "100",  5000, True))

'*********************************************************************************************************************









