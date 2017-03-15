'Get values from the DataTable
strdCIP = DataTable.Value("dCIP", dtGlobalSheet)
strUser = DataTable.Value("dCUser", dtGlobalSheet)
strPassword = DataTable.Value("dCPassword", dtGlobalSheet)

'Get a device name to install apps
intIterations = CInt(DataTable.Value("Iterations", dtGlobalSheet))
strCLIPath = Chr(34) & Environment("TestDir") & "\CLI\MobileLabs.DeviceConnect.Cli.exe" & Chr(34)
strCLIUser = DataTable.Value("CLIUser", dtGlobalSheet)
strCLIPassword = DataTable.Value("CLIPassword", dtGlobalSheet)

GetAvailableDeviceName strCLIPath, strdCIP, strCLIUser, strCLIPassword, strDeviceName, strDeviceOS
intCount = 1

'Retain the device and install apps, repeat intIterations times
Do While intCount <= intIterations

	Dim arrApp
	
	'Install these apps
	Select Case LCase(strDeviceOS)
		Case "ios"
			arrApp = Array("PhoneLookup","Wells Fargo","Good_For_Testing","Capsule","deviceControl","Trust Browser","UICatalog")
		Case "android"
			arrApp = Array("Phone Lookup","Capsule","CÜR Music_102","FirstBank","Amazon","Regions UAT3","Trust Browser")
	End Select	
	
	'Install the apps one by one
	For j = 0 To UBound(arrApp)
	
		'Release and then Retain the device first
		SetResult "Status", "Releasing device: " & strDeviceName
	
		strConnectParamRelease = strdCIP & " " & strCLIUser & " " & strCLIPassword & " -d "_
		& Chr(34) & strDeviceName & Chr(34) & " -release"
		Wait(1)
		SystemUtil.Run strCLIPath, strConnectParamRelease
		WaitForProcess "MobileLabs.DeviceConnect.Cli.exe", 60
		
		SetResult "Status", "Device released: " & strDeviceName
		
		SetResult "Status", "Retaining device: " & strDeviceName
	
		strConnectParamRetain = strdCIP & " " & strCLIUser & " " & strCLIPassword & " -d "_
		& Chr(34) & strDeviceName & Chr(34) & " -retain"
		Wait(1)
		SystemUtil.Run strCLIPath, strConnectParamRetain
		WaitForProcess "MobileLabs.DeviceConnect.Cli.exe", 60
		
		SetResult "Status", "Device retained: " & strDeviceName
	
		SetResult "Status", "Installing app: " & arrApp(j) & " on device: " & strDeviceName
		
		strConnectParam = strdCIP & " " & strCLIUser & " " & strCLIPassword & " -d "_
		& Chr(34) & strDeviceName & Chr(34) & " -install " & Chr(34) & arrApp(j) & Chr(34)
		Wait(1)
		SystemUtil.Run strCLIPath, strConnectParam
		
		blnResult = WaitForProcess("MobileLabs.DeviceConnect.Cli.exe", 90)
		
		If blnResult Then
			SetResult "Status", "App: " & arrApp(j) & " installed on device: " & strDeviceName
		Else
			SetResult "Status", "App: " & arrApp(j) & " Failed to install on device: " & strDeviceName & " and CLI.exe still running. Killing CLI.exe now!!"
			SystemUtil.CloseProcessByName "MobileLabs.DeviceConnect.Cli.exe"
		End If
		
	Next
	
	'Release the device
	SetResult "Status", "Releasing retained device: " & strDeviceName
	strConnectParam = strdCIP & " " & strCLIUser & " " & strCLIPassword & " -d "_
	& Chr(34) & strDeviceName & Chr(34) & " -release"
	SystemUtil.Run strCLIPath, strConnectParam
	
	WaitForProcess "MobileLabs.DeviceConnect.Cli.exe", 60
	SetResult "Status", "Device released: " & strDeviceName
	
	intCount = intCount+1
Loop
