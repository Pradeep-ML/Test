On Error Resume Next
'Create an object of WMI
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 

'Executing query to get the list of all wscript.exe processes
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process where name like 'QTPro.exe'")

strComputer = "."
Set objWMIService = GetObject("Winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_DesktopMonitor")
For Each objItem in colItems
	intHorizontal = objItem.ScreenWidth
	intVertical = objItem.ScreenHeight
Next

If colProcess.Count > 0 Then

	Do While colProcess.Count = 1
		Set objExplorer = CreateObject("InternetExplorer.Application")

		objExplorer.Navigate "about:blank"   
		objExplorer.ToolBar = 0
		objExplorer.StatusBar = 0
		objExplorer.Left = 22
		objExplorer.Top = intVertical - 150
		objExplorer.Width = 350
		objExplorer.Height = 50
		objExplorer.Visible = 1             
		objExplorer.Document.Body.Style.Cursor = "wait"
		
		objExplorer.Document.Title = "Mobile Labs Automation: In Progress"
		objExplorer.Document.Body.InnerHTML = "Automation Test in progress. This might take several minutes to complete."
		objExplorer.Document.Body.Style.Cursor = "default"
		Wscript.Sleep 2000
		objExplorer.Quit
		
		'Create an object of WMI
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 

		'Executing query to get the list of all wscript.exe processes
		Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process where name like 'QTPro.exe'")
	Loop

	Set objExplorer = CreateObject("InternetExplorer.Application")

	objExplorer.Navigate "about:blank"   
	objExplorer.ToolBar = 0
	objExplorer.StatusBar = 0
	objExplorer.Left = 22
	objExplorer.Top = intVertical - 150
	objExplorer.Width = 350
	objExplorer.Height = 50
	objExplorer.Visible = 1             
	objExplorer.Document.Body.Style.Cursor = "wait"
		
	objExplorer.Document.Title = "Mobile Labs Automation: Completed"
	objExplorer.Document.Body.InnerHTML = "Automation Test completed..!!"
	objExplorer.Document.Body.Style.Cursor = "default"
	Wscript.Sleep 3000
	objExplorer.Quit

End If