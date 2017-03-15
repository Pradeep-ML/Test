'
''Login to Phonelookup app
''Set environment variables from ALM
'Environment("platform") = QCUtil.CurrentTestSet.Field("CY_USER_01")
'Environment("osversion") = QCUtil.CurrentTestSet.Field("CY_USER_02")
'Environment("devicemodel") = QCUtil.CurrentTestSet.Field("CY_USER_03")
'Environment("appid") = QCUtil.CurrentTestSet.Field("CY_USER_04")
'Environment("buildnumber") = QCUtil.CurrentTestSet.Field("CY_USER_05")
'Environment("protocolversion") = QCUtil.CurrentTestSet.Field("CY_USER_06")
'Environment("agentversion") = QCUtil.CurrentTestSet.Field("CY_USER_07")
'Environment("dcip") = QCUtil.CurrentTestSet.Field("CY_USER_08")
'Environment("deviceorientation") = QCUtil.CurrentTestSet.Field("CY_USER_09")
'Environment("devicescale") = QCUtil.CurrentTestSet.Field("CY_USER_10")
'Environment("dcuser") = QCUtil.CurrentTestSet.Field("CY_USER_11")
'Environment("dcpassword") = QCUtil.CurrentTestSet.Field("CY_USER_12")
'
'
'''product version
'strProductVersion = Environment("buildnumber") 
'
'   'Fetch product version from build
'	arrValue =Split(Environment("buildnumber") , ".")
'	Environment("strProductVersion") = arrValue(0) & "." &arrValue(1)
'
'	'Fetch Build type
'	If  arrValue(2) > 600 Then
'		Environment("strBuildType") = "Nightly"	
'	Else
'		Environment("strBuildType")  = "Production"
'	End If
'
''
''	    'Get  IP address
'strIP = ""
'Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
'Set colNICs = objWMI.ExecQuery("Select * From Win32_NetworkAdapter WHERE NetConnectionID LIKE 'Local Area Connection'")
''
'For each objNIC in colNICs
'	Set colNICcfg = objWMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where MACAddress = '" & objNIC.MACAddress & "' AND IPEnabled = 'true'")
'For each objItem in colNICcfg
'	Environment("strIP") = objItem.IPAddress(0)
'Next
'Next
''
'Environment("strProduct") = "Trust"
''
''
'''Download all required files to %temp%\MobileLabsAutomation
'DownloadQCAttachments
''
'''Launch Trust Ai for the device set in ALM
'LaunchAiDisplay

'SystemUtil.Run  

On error resume next
'Login to PhoneLookup app
Login "mobilelabs" , "demo"

'Select all checkbox on search page

MobiDevice("Phone Lookup").MobiElement("Element").MobiCheckbox("Android").Set eCHECKED
wait 1
MobiDevice("Phone Lookup").MobiElement("Element").MobiCheckbox("BlackBerry").Set eCHECKED
wait 1
MobiDevice("Phone Lookup").MobiElement("Element").MobiCheckbox("iOS").Set eCHECKED
wait 1
MobiDevice("Phone Lookup").MobiElement("Element").MobiCheckbox("Windows").Set eCHECKED
wait 1
MobiDevice("Phone Lookup").MobiElement("Element").MobiDropdown("HTC").Select "Any"
wait 1
MobiDevice("Phone Lookup").MobiElement("Element").MobiRadio("All").Set 
wait 1
MobiDevice("Phone Lookup").MobiElement("Element").MobiButton("Search").Click
wait 3
'
MobiDevice("Phone Lookup").ButtonPress eBACK
wait 4

 
MobiDevice("Phone Lookup").ButtonPress eMENU @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_
wait 2
MobiDevice("Phone Lookup").MobiList("List").Select "Controls" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf2.xml_;_
wait 2

MobiDevice("Phone Lookup").MobiList("List_2").Scroll eBOTTOM
wait 3
MobiDevice("Phone Lookup").MobiList("List_2").Scroll eTOP
wait 3

MobiDevice("Phone Lookup").MobiList("List_2").Select "DatePicker"
wait 1

LogOut



'##########################################################################################################################
''@Sub:        		LaunchAiDisplay
''@Description: 	This sub connects to the device setup in ALM TestSet fields. If the device can't be connected to then
''					a message box will be displayed which will ask the user to connect manually and then continue
''@Created By: 		Naveen 
''@Created On: 		05/28/2013
''@Last Updated:	08/31/2013
''--------------------------------------------------------------------------------------------------------------------------
''@Example: 		LaunchAiDisplay
''--------------------------------------------------------------------------------------------------------------------------

Sub LaunchAiDisplay

	'If Trust Ai is already running then verify that it is the correct OS, osversion and Device Model
	Set objDevice = MobiDevice("Class Name:=MobiDevice")
	If objDevice.Exist(2) Then
		strDeviceType = objDevice.GetROProperty("devicetype")
		strOSVersion = objDevice.GetROProperty("osversion")
		strPlatform = objDevice.GetROProperty("platform")
		If InStr(1, Environment("devicemodel"), strDeviceType, 1) > 0 _ 
		AND InStr(1, strOSVersion, Environment("osversion"), 1) > 0 _
		AND InStr(1, strPlatform, Environment("platform"), 1) > 0 Then
			Exit Sub
		End If	
	End If
	SystemUtil.CloseProcessByName "MobileLabs.deviceViewer.exe"
	Wait 1	
	
	'Make sure all values are defined in QC TestSet
	blnInfo = False
	If IsEmpty(Environment("platform")) Then
		strVars = strVars & "platform, "
	ElseIf IsEmpty(Environment("osversion")) Then
		strVars = strVars & "osversion, "
	ElseIf IsEmpty(Environment("devicemodel")) Then
		strVars = strVars & "devicemodel, "
	ElseIf IsEmpty(Environment("appid")) Then
		strVars = strVars & "appid, "
	ElseIf IsEmpty(Environment("buildnumber")) Then
		strVars = strVars & "buildnumber, "
	ElseIf IsEmpty(Environment("protocolversion")) Then
		strVars = strVars & "protocolversion, "
	ElseIf IsEmpty(Environment("agentversion")) Then
		strVars = strVars & "agentversion, "
	ElseIf IsEmpty(Environment("dcip")) Then
		strVars = strVars & "dcip, "
	ElseIf IsEmpty(Environment("deviceorientation")) Then
		strVars = strVars & "deviceorientation, "
	ElseIf IsEmpty(Environment("devicescale")) Then
		strVars = strVars & "devicescale, "
	Else
		blnInfo = True
	End If
	
	If Not(blnInfo) Then
		MsgBox "One of these fields don't have a value defined in ALM TestSet: " & strVars & VBNewLine _
		& "Please add these values and run the Testset again!"
		ExitTest
	End If

	'Add cases for platforms here. Android and iOS
	If LCase(Environment("platform")) = "androidos" Then
		Environment("platform") = 1
	ElseIf LCase(Environment("platform")) = "iphone os" Then
		Environment("platform") = 0
	End If
			
	strPassSalt = "pass-56ffc8c3cd680bf96ba600943f149b92"
	strConnectionString = "Server=localhost;Port=5433;Database=deviceconnect_app;User Id=" &_
	"deviceconnect_app;Password=" & strPassSalt
	strQuery = "select id from device where friendly_model='" & Environment("devicemodel") &_
	"' AND operating_system_version='" & Environment("osversion") & "' AND operating_system=" &_
	Environment("platform") & " AND availability=2;"
		
	strGUID = FetchGUID(Environment("dcip"), strConnectionString, strQuery)
	
	'If an empty GUID is returned then ask the user to manually connect to the device to continue
	If IsEmpty(strGUID) Then
		MsgBox "No DeviceID could be fetched. Please connect to the device manually via dC UI and hit OK to continue!"
		Exit Sub
	End If
			
	'Launch the Device Controller as setup in the ControllerConfig.xlsx
	SystemUtil.Run GetTrustInstallDir & "MobileLabs.deviceViewer.exe", _ 
	"-url airstream://launchApp?" _ 
	& "hubAddress=" & Environment("dcip") & ":10160" _ 
	& "&deviceId=" & strGUID _ 
	& "&applicationId=" & Environment("appid") _ 
	& "&username=" & Environment("dcuser") _ 
	& "&password=" & Environment("dcpassword") _ 
	& "&autoconnect=true" _
	& "&install=true" & " -scale " & Environment("devicescale") _
	& " -orientation " & Environment("deviceorientation")
	
	
	'Check if the MobileLabs.Trust.AiDisplay process is running or not. Wait for 30 seconds max
	intCount = 0
	blnTrustAi = False

	Do While intCount < 30
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
		Set colProcess = objWMIService.ExecQuery _
		("Select * from Win32_Process where name like 'MobileLabs.deviceViewer.exe'")
		If colProcess.Count > 0 Then
			blnTrustAi = True
			Set colProcess = Nothing
			Set objWMIService = Nothing
			Wait(10)
			Exit Do
		End If

		Set colProcess = Nothing
		Set objWMIService = Nothing
		Wait(1)
		intCount = intCount + 1
	Loop

	On Error Resume Next
    intExitCount = 1
	Do While IsEmpty(MobiDevice("micclass:=MobiDevice").GetROProperty("platform"))
		intExitCount = intExitCount + 1
		If intExitCount > 10 Then 'Wait max 200 seconds for Trust Ai to launch device screen
			Exit Do
		End If
	Loop
	On Error GoTo 0

	'If the device couldn't be connected then report a failure and exit test
	If Not(MobiDevice("micclass:=MobiDevice").Exist) Then
		Reporter.ReportEvent micFail, "Launch Trust Ai Display", "Failed to make connection: " & VbNewLine &_
		"Device Model: " & strDeviceModel &_
		"Device OS and OS Version: " & strOS & ", " & strOSVersion &_
		"GUID: " & strGUID &_
		"App: " & strApp
		ExitTest
	End If

End Sub


'##########################################################################################################################
''@Function:        FetchGUID
''@Description: 	Connects to postgres db and gets back the GUID for a device
''@Created By: 		Naveen
''@Created On: 		08/19/2012
''--------------------------------------------------------------------------------------------------------------------------
''@Param Name:  	strHUB 
''@Param Type: 		String
''@Param Drtn: 		In
''@Param Desc: 		IP of the dC box
''--------------------------------------------------------------------------------------------------------------------------
''@Param Name:  	strConnectionString 
''@Param Type: 		String
''@Param Drtn: 		In
''@Param Desc: 		Connection string to establish connection with postgres db
''--------------------------------------------------------------------------------------------------------------------------
''@Param Name: 		strQuery
''@Param Type: 		String
''@Param Drtn: 		In
''@Param Desc: 		db query to be executed
''--------------------------------------------------------------------------------------------------------------------------
''@Example:   	strConnectionString = "Server=localhost;Port=5433;Database=deviceconnect_app;User Id=" &_
''				"deviceconnect_app;Password=" & strPassSalt
''				strQuery = "select id from device where friendly_model='" & strModel &_
''				"' AND operating_system_version='" & strOSVersion & "';"	
''				strGUID = FetchGUID("10.4.1.27", "Apple - iPhone 4", _
''				"6.1.3",strConnectionString, strQuery)
''--------------------------------------------------------------------------------------------------------------------------

Function FetchGUID(strHUB, strConnectionString, strQuery)
	SystemUtil.CloseProcessByName "cmd.exe"
	SystemUtil.Run "cmd.exe"
	
	strStartPlink = """" & GetFilePath("plink.exe") & """" & " -L 5433:127.0.0.1:5432 -ssh -P 22 -2 -C -l deviceconnect -pw GoMobile! " & strHUB
	
	Window("regexpwndclass:=ConsoleWindowClass").Type "cd\"
	Window("regexpwndclass:=ConsoleWindowClass").Type micReturn
	Wait 2
	Window("regexpwndclass:=ConsoleWindowClass").Type strStartPlink
	Window("regexpwndclass:=ConsoleWindowClass").Type micReturn
	
	Wait 25
		
	Set GetGUID = DotNetFactory.CreateInstance("FetchDeviceGUID.GetGUID", GetFilePath("FetchDeviceGUID.dll"))
	FetchGUID = GetGUID.ExecuteQuery(strConnectionString, strQuery)
	SystemUtil.CloseProcessByName "cmd.exe"
End Function


'##########################################################################################################################
''@Function:  Login
''@Description:  Login to App
''@Return Type:	Boolean
''@Created By: Amit
''@Created On: 06/19/2013
''--------------------------------------------------------------------------------------------------------------------------
''@Param Name:  strUserName 
''@Param Type:   String
''@Param Drtn: 		In
''@Param Desc: 	UserName for login to app
''--------------------------------------------------------------------------------------------------------------------------
''@Param Name:  strPassword 
''@Param Type:   String
''@Param Drtn: 		In
''@Param Desc: 	Password for login to app
''--------------------------------------------------------------------------------------------------------------------------
''@Example: blnFlag = Login("mobilelabs", "demo")
'--------------------------------------------------------------------------------------------------------------------------
Function Login(strUserName, strPassword)

	'Call the function  Logout
	Logout
	Set objDevice = MobiDevice("name:=Phone.*")
	strPlatform = objDevice.getroproperty("platform")
	If Instr(1 , LCase(strPlatform), "iphone")  > 0 Then
		'Set  Username and Password
		objDevice.MobiEdit("id:=2").Set(strUserName)
		objDevice.MobiEdit("id:=3").Set(strPassword)
		objDevice.MobiSwitch("name:=Switch").Set eACTIVATE

	Else
		'Set  Username and Password
		objDevice.MobiEdit("defaultvalue:= Enter Username").Set(strUserName)
		objDevice.MobiEdit("defaultvalue:= Enter Password").Set(strPassword)
	End If
	wait 2
	'Click on Sign In button
	objDevice.MobiButton("name:=SignIn").Click
	
	wait(7)
	If objDevice.MobiElement("text:=Search","nativeclass:=UINavigationItemView").Exist(3)  Then
		Login = True
	Else
		Login = False
	End If
End Function

''##########################################################################################################################
'''@Sub:  Logout
'''@Description:  Logs out of the PhoneLookup app
'''@Return Type:	Boolean
'''@Created By: Naveen
'''@Created On: 01/25/2013
'''--------------------------------------------------------------------------------------------------------------------------
'''@Example: Logout
''--------------------------------------------------------------------------------------------------------------------------
'Sub Logout
'	 'create description for MobiDevice
'	Set objDevice = Description.Create
'	objDevice("name").Value = "PhoneLookup"
'
'	'create description for MobiButton
'		Set objButton = Description.Create
'		objButton("micclass").Value = "MobiButton"
'	
'	'Loop until Base state is reached
'	Do Until MobiDevice(objDevice).MobiButton("name:=SignIn").Exist(2)
'		Set btnObjects =  MobiDevice(objDevice).ChildObjects(objButton)
'	
'		For  i = 0  to  btnObjects.Count-1
'		
'			If btnObjects(i).GetROProperty("name") = "Button" OR btnObjects(i).GetROProperty("name") = "Logout" Then
'				'Click on button
'				btnObjects(i).Click
'				Exit For
'			End If
'		
'		Next
'		Set btnObjects = Nothing
'	Loop
'
'	Set objButton = Nothing
'	Set objDevice = Nothing
'End Sub



'##########################################################################################################################
''@Function:  LogOut
''@Description:  LogOut and navigate to Login Screen (Applicable for both Android and IOS Platform)
''@Return Type:	Boolean
''@Created By: Shweta
''@Created On: 30/01/2013
''--------------------------------------------------------------------------------------------------------------------------
''@Example:
'flag =  LogOut
'--------------------------------------------------------------------------------------------------------------------------

Function LogOut 

	On Error Resume Next

	'Setting initial return value
	LogOut  =  False

	'MobiDevice Description
   Set objDevice = Description.Create
   objDevice("micclass").Value = "MobiDevice"

	'Fetching platform value at run time
   strPlatform  = MobiDevice(objDevice).GetROProperty("platform")

	If Not MobiDevice(objDevice).MobiElement("text:=Username").Exist(3) Then

		'Android Platform
		If  Instr(1 , LCase(strPlatform), "android")  > 0 Then
		
			'Open Menu
			MobiDevice(objDevice).ButtonPress eMENU
			
			'Verify existence of Logout Element in Menu 
			If MobiDevice(objDevice).MobiElement("name:=LogOut").Exist(5)   Then
				'Click Logout Element
				MobiDevice(objDevice).MobiElement("name:=LogOut").Click
				'Refresh Object
				MobiDevice(objDevice).RefreshObject
			Else
				Do While Not  MobiDevice(objDevice).MobiElement("name:=LogOut").Exist  
				'Click on back button in case menu option is not available
				MobiDevice(objDevice).ButtonPress eBACK
				
				Wait 3
				'Open menu
				MobiDevice(objDevice).ButtonPress eMENU
				wait 4
				Loop
				MobiDevice(objDevice).MobiElement("name:=LogOut").Click
				MobiDevice(objDevice).RefreshObject
			End If
		
		'IOS Platform
		ElseIf  Instr(1 , LCase(strPlatform), "iphone")  > 0  Then
		
			'create description for MobiButton
			'	Set objButton = Description.Create
			'	objButton("micclass").Value = "MobiButton"
			
			'Loop until Login screen is reached
			Do Until MobiDevice(objDevice).MobiElement("text:=Sign In").Exist(2)
				If  MobiDevice(objDevice).MobiButton("name:= Logout").Exist(1)Then
					MobiDevice(objDevice).MobiButton("name:= Logout").Click
					Exit Do 
				End If
				MobiDevice(objDevice).MobiButton("location:= 0").Click
				'	Set btnObjects =  MobiDevice(objDevice).ChildObjects(objButton)
				'
				'			For  i = 0  to  btnObjects.Count-1
				'				If btnObjects(i).GetROProperty("name") = "Button" OR btnObjects(i).GetROProperty("name") = "Logout" Then
				'					'Click on button
				'					btnObjects(i).Click
				'					wait (2)
				'				End If
				'			Next
			Loop
			
		End If 

	End If

	'Verify Existence of  Sign In element
	If  MobiDevice(objDevice).MobiElement("text:=Username").Exist(5)  Then
			Reporter.ReportEvent micPass, Environment("StepName"), "Successfully navigated to Login Screen" 
			LogOut = True
	Else
			strSummary = "Logout  failed"
			strActualResult = "Failed to navigate to Login Screen"
			Reporter.ReportEvent micFail, Environment("StepName"), strActualResult
			ExitTest
	End If

End Function
