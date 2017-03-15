'##########################################################################################################
'# Phone Lookup App Demo Script
'# Version 4.0
'# ©2013 Mobile Labs
'##########################################################################################################
'Option Explicit

Dim product
Dim repository
Dim tableIndex
Dim productIndex
Dim connectParams
Dim dcURL
Dim dcDeviceID
Dim dcAppID
Dim dcUsername
Dim dcPassword
Dim dcScales
Dim dcAutoconnect
Dim dcInstallApp
Dim i
Dim z
Dim oFSO
Dim trust1
Dim strProcess

'DownLoadCLI to Local Temp
DownloadCLI

'ALM Integration
dcURL = QCUtil.CurrentTestSetTest.Field("TC_USER_05")
dcDeviceID = QCUtil.CurrentTestSetTest.Field("TC_USER_01")
dcAppID = QCUtil.CurrentTestSetTest.Field("TC_USER_02")
dcUsername = QCUtil.CurrentTestSetTest.Field("TC_USER_03")
dcPassword = QCUtil.CurrentTestSetTest.Field("TC_USER_04")
dcScales = QCUtil.CurrentTestSetTest.Field("TC_USER_07")


For z= 1 To 1
	Reporter.ReportEvent micDone, "Iteration #" & z, "Step passed."
	Services.StartTransaction "TransactionTiming #"& z 


If z=1 Then
	strConnectParam = dcURL  & " " & dcUsername & " " & dcPassword &" "& "-d" & " " _
	& Chr(34) & dcDeviceID  & Chr(34) & " " & "-scale" & " " & dcScales & " " _
	& "-r" & " " & Chr(34) & dcAppID & Chr(34)  

	SystemUtil.CloseProcessByName "MobileLabs.deviceViewer.exe"
    'Launch device using CLI
	SystemUtil.Run "MobileLabs.DeviceConnect.Cli.exe", strConnectParam , Environment("SystemTempDir")&"\CLI"	
	Wait(5)
	Reporter.ReportEvent micDone, "ConnectToDevice", "Connect Parameters: " & strConnectParam
	'SystemUtil.Run "MobileLabs.DeviceConnect.Cli.exe", strConnectParam , Environment("SystemTempDir") & "\CLI"
End If


'Check for Login Screen 
If MobiDevice("PhoneLookup").MobiEdit("Username").Exist(300) =  False Then
    Reporter.ReportEvent micFail, "Login", "Failed to Login, the Login Screen is not displaying"
	ExitTest
End If

'Demo Scale
MobiDevice("PhoneLookup").Scale 100
Wait 3
MobiDevice("PhoneLookup").Scale 75
Wait 3
MobiDevice("PhoneLookup").Scale 50
Wait 3
MobiDevice("PhoneLookup").Scale 25
Wait 3
'Reset scale to what was chosen in ALM
MobiDevice("PhoneLookup").Scale CInt(dcScales)

DataTable.SetCurrentRow(1)
'Login with username/pwd and descriptive programming
MobiDevice("PhoneLookup").MobiEdit("Username").Set datatable.Value("Username", "Action1")
MobiDevice("PhoneLookup").MobiEdit("Password").Set datatable.Value("Password", "Action1")
MobiDevice("PhoneLookup").MobiButton("text:=Sign In").Click
wait(1)

'Catch error and input correct username/password
If MobiDevice("PhoneLookup").GetROProperty("platform")="AndroidOS" Then
	MobiDevice("PhoneLookup").MobiButton("OK").Click
Else
	If MobiDevice("PhoneLookup").InsightObject("InsightObject1").Exist(1) Then
		MobiDevice("PhoneLookup").InsightObject("InsightObject1").Click
	End If
	
	If MobiDevice("PhoneLookup").InsightObject("InsightObject2").Exist(1) Then
	MobiDevice("PhoneLookup").InsightObject("InsightObject2").Click
	End If
End If

DataTable.SetCurrentRow(2)
MobiDevice("PhoneLookup").MobiEdit("Username").Set datatable.Value("Username", "Action1")
MobiDevice("PhoneLookup").MobiEdit("Password").Set datatable.Value("Password", "Action1")
MobiDevice("PhoneLookup").MobiButton("text:=Sign In").Click

'In case keyboard is up on Android, minimize it and reinsert credentials
If (MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS") Then
	If MobiDevice("PhoneLookup").MobiButton("Search").Exist(5) = False Then
	    MobiDevice("PhoneLookup").ButtonPress eBACK
	    MobiDevice("PhoneLooku").MobiEdit("Username").Set "mobilelabs"
	    MobiDevice("PhoneLookup").MobiEdit("Password").Set "demo"
	    MobiDevice("PhoneLookup").MobiButton("text:=Sign In").Click
    End If
End If

'Navigate to Controls
selectMenuOption "Controls"

'Demo Rotate
MobiDevice("PhoneLookup").Rotate eLANDSCAPELEFT
Wait 3
MobiDevice("PhoneLookup").Rotate eLANDSCAPERIGHT
Wait 3
MobiDevice("PhoneLookup").Rotate ePORTRAIT
'UpsideDown orientation is not supported on most phones and therefore is not demostrated here.

'Showcase swipe
'Use Web View for Android or MapView for iOS
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    MobiDevice("PhoneLookup").MobiList("Controls").Select "ZoomControls"
	
	MobiDevice("PhoneLookup").MobiEdit("httpwwwgooglemapscom").WaitProperty "Visible", True, 5000
	MobiDevice("PhoneLookup").MobiEdit("httpwwwgooglemapscom").Set "http://www.googlemaps.com"
	MobiDevice("PhoneLookup").ButtonPress eBACK 
	
	'Hide keyboard
	MobiDevice("PhoneLookup").Type vbCr
	Wait (6)	
Else
   MobiDevice("PhoneLookup").MobiList("ControlsList").Select "MKMapView"
   MobiDevice("PhoneLookup").MobiElement("MKMapView").WaitProperty "Visible", True, 5000
   	If MobiDevice("PhoneLookup").InsightObject("InsightObject1").Exist(2) Then
		MobiDevice("PhoneLookup").InsightObject("InsightObject1").Click
	End If
End If

MobiDevice("PhoneLookup").Swipe eRIGHT, eMEDIUM, 40, 70
Wait 3
MobiDevice("PhoneLookup").Swipe eDOWN, eMEDIUM, 40, 70
Wait 3
MobiDevice("PhoneLookup").Swipe eLEFT, eMEDIUM, 40, 70
Wait 3
MobiDevice("PhoneLookup").Swipe eUP, eMEDIUM, 40, 70
Wait 3

'Demo Draw
'Swipe Down
MobiDevice("PhoneLookup").Draw "down(50%, 70%) move(50%, 40%) up()"
MobiDevice("PhoneLookup").Draw "sync()"
'Swipe Up
MobiDevice("PhoneLookup").Draw "down(50%, 40%) move(50%, 70%) up()"
MobiDevice("PhoneLookup").Draw "sync()"
'Swipe Right
MobiDevice("PhoneLookup").Draw "down(70%, 50%) move(40%, 50%) up()"
MobiDevice("PhoneLookup").Draw "sync()"
'Swipe Left
MobiDevice("PhoneLookup").Draw "down(40%, 50%) move(70%, 50%) up()"
MobiDevice("PhoneLookup").Draw "sync()"
'Draw a line an arc around then another line
MobiDevice("PhoneLookup").Draw "down(50%, 50%) move(50%, 60%) arc(50%, 50%, 90, duration=5s) move(70%, 70%) up()"
MobiDevice("PhoneLookup").Draw "sync()"

'Select which Controls for iOS
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    MobiDevice("PhoneLookup").ButtonPress eBACK 
    
Else		
    MobiDevice("PhoneLookup").MobiButton("Controls").Click
    
    If Not(MobiDevice("PhoneLookup").MobiList("ControlsList").Exist(3)) Then
        MobiDevice("PhoneLookup").MobiButton("Controls").Click
    End If
End If
	
'Showcase Hybrid
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    MobiDevice("PhoneLookup").MobiList("Controls").WaitProperty "Visible", True, 5000
    MobiDevice("PhoneLookup").MobiList("Controls").Select "ZoomControls"
Else
    MobiDevice("PhoneLookup").MobiList("ControlsList").Select "UIWebView"
    MobiDevice("PhoneLookup").MobiEdit("httpwwwmobilelabsinccom").WaitProperty "Visible", True, 5000
    MobiDevice("PhoneLookup").MobiEdit("httpwwwmobilelabsinccom").Set "http://www.yelp.com"
    MobiDevice("PhoneLookup").MobiEdit("httpwwwmobilelabsinccom").Click
    
    MobiDevice("PhoneLookup").Type vbCr 
    wait (2)
    If Not(MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebElement("YelpLogo").Exist(2)=True) Then
	   MobiDevice("PhoneLookup").MobiEdit("httpwwwmobilelabsinccom").Click
	   MobiDevice("PhoneLookup").Type vbCr 
	   wait(2)
	End If
End If 

'Goto Yelp and highlight objects
If MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebElement("YelpLogo").Exist(10) Then
    If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
	   MobiDevice("PhoneLookup").ButtonPress eBACK
	End If
	MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebElement("YelpLogo").highlight
	MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebElement("RestaurantsLink").highlight
	'MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebEdit("SearchYelp").Set "Fajitas"
	
	MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebDropdown("ChangeLanguage").ScrollIntoView
	Wait 2
	MobiDevice("PhoneLookup").MobiWebView("WebView").MobiWebDropdown("ChangeLanguage").Click
					
	'Close dropdown
	If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
	   If MobiDevice("PhoneLookup").GetROProperty("devicetype") = "GT-N8013" Then
	       If MobiDevice("PhoneLookup").MobiButton("Done").Exist(5) Then
    	       MobiDevice("PhoneLookup").MobiButton("Done").Click
    	   End If
	   Else
	       If MobiDevice("PhoneLookup").MobiElement("ChangeLanguage").Exist(14) Then
	           MobiDevice("PhoneLookup").MobiElement("ChangeLanguage").Click
           End If
       End If
    Else
        MobiDevice("PhoneLookup").MobiButton("Done").Click
    End If
End If
	
'Leave Hybrid page
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    wait(2)
    MobiDevice("PhoneLookup").ButtonPress eBACK
Else
	MobiDevice("PhoneLookup").MobiButton("Controls").Click
End If

If NOT MobiDevice("PhoneLookup").MobiElement("text:=Controls").Exist(5)  AND MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS"  Then
	MobiDevice("PhoneLookup").ButtonPress eBACK
End If
'Select Search
selectMenuOption "Search"

If MobiDevice("PhoneLookup").MobiButton("Search").Exist(10) = False Then
    Reporter.ReportEvent micFail, "Search", "Failed to Search, the Search Screen is not displaying" 
End If
	
'Select 'Any' from dropdown
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    MobiDevice("PhoneLookup").MobiDropdown("Manufacturer").WaitProperty "Visible", True, 2000
    MobiDevice("PhoneLookup").MobiDropdown("Manufacturer").Select "Any"
Else
    wait(2)
	MobiDevice("PhoneLookup").MobiEdit("Manufacturer").WaitProperty "Visible", True, 5000
    MobiDevice("PhoneLookup").MobiEdit("Manufacturer").Click
    wait(2)
    MobiDevice("PhoneLookup").MobiPicker("Picker").WaitProperty "Visible", True, 5000
    MobiDevice("PhoneLookup").MobiPicker("Picker").Select "Any"
    MobiDevice("PhoneLookup").MobiButton("Done").Click	
End If
	
'Set operating system and inventory
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    MobiDevice("PhoneLookup").MobiElement("OperatingSystem").MobiCheckbox("iOS").Set eCHECKED
	MobiDevice("PhoneLookup").MobiElement("OperatingSystem").MobiCheckbox("Android").Set eCHECKED
	MobiDevice("PhoneLookup").MobiElement("OperatingSystem").MobiCheckbox("BlackBerry").Set eCHECKED
	MobiDevice("PhoneLookup").MobiElement("OperatingSystem").MobiCheckbox("Windows").Set eCHECKED
	
	MobiDevice("PhoneLookup").MobiElement("OperatingSystem").MobiRadio("All").Set
Else
	MobiDevice("PhoneLookup").MobiSwitch("AndroidSwitch").Set eACTIVATE
	MobiDevice("PhoneLookup").MobiSwitch("BlackBerrySwitch").Set eACTIVATE
	MobiDevice("PhoneLookup").MobiSwitch("iOSSwitch").Set eACTIVATE
	MobiDevice("PhoneLookup").MobiSwitch("WindowsSwitch").Set eACTIVATE
		
	MobiDevice("PhoneLookup").MobiSegment("InventorySegment").Select 0
End If
	
'Click Search
MobiDevice("PhoneLookup").MobiButton("Search").Click
	
'Verify Search List came up
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
    If MobiDevice("PhoneLookup").MobiList("andSearchResults").Exist(5) = False Then
	    If MobiDevice("PhoneLookup").MobiElement("OperatingSystem").MobiCheckbox("Android").Exist(2) Then
	        MobiDevice("PhoneLookup").MobiButton("Search").Click
		End If
	End If	
Else
    If MobiDevice("PhoneLookup").MobiList("iosSearchResults").Exist(5) = False Then
		If MobiDevice("PhoneLookup").MobiSwitch("AndroidSwitch").Exist(2) Then
		    MobiDevice("PhoneLookup").MobiButton("Search").Click
		End If
	End If
End If
	
'Select 2 Products and output onscreen data to UFT Global datasheet
tableIndex = 0
For tableIndex = 1 To 2
   	wait(4)
	If tableIndex = 1 Then
		product = "iPad 2"
	Else
		product = "MC75A"
	End If
		
	If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
		productIndex = 0
		MobiDevice("PhoneLookup").MobiList("andSearchResults").WaitProperty "Visible", True, 5000
		MobiDevice("PhoneLookup").MobiList("andSearchResults").Select product
	Else
		productIndex = 1
		MobiDevice("PhoneLookup").MobiList("iosSearchResults").WaitProperty "Visible", True, 5000
		MobiDevice("PhoneLookup").MobiList("iosSearchResults").Select product
	End If
	
	'Lets add the onscreen data to QTP spreadsheet
	addProductToDataTable product, tableIndex, productIndex 
	
	'Checkpoint
	If tableIndex = 1 Then
		MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("Price").Check CheckPoint("Price49999Each")
	End If
	
	'Return to Search screen
	If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
		MobiDevice("PhoneLookup").ButtonPress eBACK 
	Else
        MobiDevice("PhoneLookup").MobiButton("Results").WaitProperty "Visible", True, 5000
		MobiDevice("PhoneLookup").MobiButton("Results").Click
	End If
Next
	
'Verify we returned to proper Search screen
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
	MobiDevice("PhoneLookup").MobiList("andSearchResults").WaitProperty "Visible", True, 5000
	MobiDevice("PhoneLookup").ButtonPress eBACK 
Else
    MobiDevice("PhoneLookup").MobiButton("SearchBack").WaitProperty "Visible", True, 5000
	MobiDevice("PhoneLookup").MobiButton("SearchBack").Click
End If
	
'Log out of PhoneLookup application
If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
	selectMenuOption "LogOut"
Else
    MobiDevice("PhoneLookup").MobiButton("Logout").WaitProperty "Visible", True, 5000
	MobiDevice("PhoneLookup").MobiButton("Logout").Click
End If

Wait (2)

If z=1 Then
	
	'Close AiDisplay
	SystemUtil.CloseProcessByName  "MobileLabs.deviceViewer.exe"

End If

Services.EndTransaction "TransactionTiming #"& z
Next

Sub selectMenuOption(ByVal menu)
    If MobiDevice("PhoneLookup").GetROProperty("platform") = "AndroidOS" Then
	    MobiDevice("PhoneLookup").ButtonPress eMENU
	    Wait 2
	
	    MobiDevice("PhoneLookup").MobiElement(menu).WaitProperty "Visible", True, 5000
	    MobiDevice("PhoneLookup").MobiElement(menu).Click
	Else
	    MobiDevice("PhoneLookup").MobiButton("Toolbar" & menu).WaitProperty "Visible", True, 5000
	    MobiDevice("PhoneLookup").MobiButton("Toolbar" & menu).Click
	End If
End Sub

Sub addProductToDataTable(ByVal productSearch, ByVal tableRow, ByVal productIndex)
	Dim productName
	Dim productSKU
	Dim productOS
	Dim productMan
	Dim productPrice
	Dim productQTY
	Dim skuArray
	Dim osArray
	Dim manArray
	Dim priceArray
	Dim qtyArray

	DataTable.SetCurrentRow(tableRow)

	productName = MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("Text:=" & productSearch, "Visible:=True", "Index:=0").GetROProperty("text")
	datatable.Value("Name", "Global") = productName

	productSKU = MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("ProductSKU" & productIndex).GetROProperty("text")
	skuArray = Split(productSKU, " # ", -1, 1)
	datatable.Value("SKU", "Global") = skuArray(productIndex)

	productOS = MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("OperatingSystem" & productIndex).GetROProperty("text")
	osArray = Split(productOS, " : ", -1, 1)
	datatable.Value("OS", "Global") = osArray(productIndex)

	productMan = MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("Manufacturer" & productIndex).GetROProperty("text")
	manArray = Split(productMan, " : ", -1, 1)
	datatable.Value("Manufacturer", "Global") = manArray(productIndex)
		
	productPrice = MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("Price").GetROProperty("text")
	priceArray = Split(productPrice, " : ", -1, 1)
	datatable.Value("Price", "Global") = priceArray(productIndex)
		
	productQTY = MobiDevice("PhoneLookup").MobiElement("ProductDetails").MobiElement("OnlineQuantity" & productIndex).GetROProperty("text")
	qtyArray = Split(productQTY, " : ", -1, 1)
	datatable.Value("Quantity", "Global") = qtyArray(productIndex)
End Sub


'############################



Sub  DownloadCLI
	strTempPath = Environment("SystemTempDir") & "\CLI"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(strTempPath) Then
		objFSO.DeleteFolder strTempPath
		Wait(2)
	End If
	objFSO.CreateFolder strTempPath

	strQCDir = "Subject\Trust\Framework\Files"
	
	Set objFolder = QCUtil.QCConnection.TreeManager.NodeByPath(strQCDir) 
	Set objAttachmentList = objFolder.Attachments.NewList("") 
	
	For Each objAttachment In objAttachmentList 
		Set objExtStorage = objAttachment.AttachmentStorage 
		objAttachmentName = objAttachment.DirectLink
		objExtStorage.Load objAttachmentName, true 
		objFSO.MoveFile objExtStorage.ClientPath & "\" & objAttachmentName, strTempPath & "\" & Split(objAttachment.Name, "_")(UBound(Split(objAttachment.Name, "_")))
	Next 
End Sub



