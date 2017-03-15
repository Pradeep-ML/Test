'##########################################################################################################################
''@Sub:				CopyToolsToTemp
''@Description::	Copy all files from %\MobileLabsAutomation\Tools and it's subfolders to SystemTempDir/MobileLabsAutomation
''@Return Type:		N/A
''@Created By: 		Naveen
''@Created On: 		10-Sept-2015
''Modified By :     
''Modified On : 	
''--------------------------------------------------------------------------------------------------------------------------
''@Example:  CopyToolsToTemp
''--------------------------------------------------------------------------------------------------------------------------

'Sub CopyToolsToTemp
'	strTempPath = Environment("SystemTempDir") & "\MobileLabsAutomation"
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	If objFSO.FolderExists(strTempPath) Then
'		objFSO.DeleteFolder strTempPath
'		Wait(2)
'	End If
'	objFSO.CreateFolder strTempPath
'	
'	strToolsPath = GetRootFolderPath & "Tools"
'	
'	'Copy all files in %\MobileLabsAutomation\Tools
'	For Each File In objFSO.GetFolder(strToolsPath).Files 
'		 objFSO.GetFile(File).Copy strTempPath & "\" & objFSO.GetFileName(File),True 
'		 Print "Copying : " & Chr(34) & objFSO.GetFile(File) & Chr(34) & " to " & strTempPath 
'	Next
'	
'	'Copy all files in all subfolders of %\MobileLabsAutomation\Tools
'	Set objToolsSubFolder = objFSO.GetFolder(strToolsPath).SubFolders
'	For Each SubFolder In objToolsSubFolder
'		For Each File1 In objFSO.GetFolder(SubFolder).Files 
'			 objFSO.GetFile(File1).Copy strTempPath & "\" & objFSO.GetFileName(File1),True 
'			 Print "Copying : " & Chr(34) & objFSO.GetFile(File1) & Chr(34) & " to " & strTempPath
'		Next
'	Next
'
'End Sub
'
'
'
'
'Function GetRootFolderPath
'	'Get current directory
'	strCurrentPath = Environment("TestDir")
'
'	'Loop until "MobileLabsAutomationFramework" folder is found
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	blnRootFolderFound = True
'	strRootPath = strCurrentPath
'	Do While Replace(Split(strRootPath,"\")(UBound(Split(strRootPath,"\"))), " ", "") <> "MobileLabsAutomationFramework"
'		strRootPath = objFSO.GetParentFolderName(strRootPath)
'		'Exit if reaches the system drive
'		If InStr(1, strRootPath, "\") = 0 Then
'			blnRootFolderFound = False
'			Exit Do
'		End If
'	Loop
'
'	'Define the path of the Root Folder: <MobileLabs Automation Framework>
'	If blnRootFolderFound Then
'		If Right(strRootPath,1) <> "\" Then
'			strRootPath = strRootPath & "\"
'		End If
'	Else
'		MsgBox "Could not find the root folder MobileLabsAutomationFramework. Please check that you are executing the test from correct location."
'		ExitTest
'	End If
'	
'	GetRootFolderPath = strRootPath
'	
'	Set objFSO = Nothing
'End Function













'ReadEnvironmentVariables
'
'Print "platform: " & Environment("platform")
'Print "osversion: " & Environment("osversion")
'Print "devicemodel: " & Environment("devicemodel")
'Print "appid: " & Environment("appid")
'Print "buildnumber: " & Environment("buildnumber")
'Print "protocolversion: " & Environment("protocolversion")
'Print "agentversion: " & Environment("agentversion")
'Print "dcip: " & Environment("dcip")
'Print "deviceorientation: " & Environment("deviceorientation")
'Print "devicescale: " & Environment("devicescale")
'Print "dcuser: " & Environment("dcuser")
'Print "dcpassword: " & Environment("dcpassword")
'Print "serveruser: " & Environment("serveruser") 
'Print "serverpassword: " & Environment("serverpassword")
'Print "dcversion: " & Environment("dcversion")

'Sub ReadEnvironmentVariables
'	
'	'Get current directory
'	strCurrentPath = Environment("TestDir")
'	
''Go to the root folder path and setup the path to the TestLab.xlsx
'
'	'Loop until "MobileLabsAutomationFramework" folder is found
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	blnRootFolderFound = True
'	strRootPath = strCurrentPath
'	Do While Replace(Split(strRootPath,"\")(UBound(Split(strRootPath,"\"))), " ", "") <> "MobileLabsAutomationFramework"
'		strRootPath = objFSO.GetParentFolderName(strRootPath)
'		'Exit if reaches the system drive
'		If InStr(1, strRootPath, "\") = 0 Then
'			blnRootFolderFound = False
'			Exit Do
'		End If
'	Loop
'
'	'Define the path of the Root Folder: <MobileLabs Automation Framework>
'	If blnRootFolderFound Then
'		If Right(strRootPath,1) <> "\" Then
'			strRootPath = strRootPath & "\"
'		End If
'	Else
'		MsgBox "Could not find the root folder MobileLabsAutomationFramework. Please check that you are executing the test from correct location."
'		ExitTest
'	End If
'	
'	strTestSetPath = strRootPath & "Environment\TestLab.xlsx"
'	
'	Set objExcel = CreateObject("Excel.Application")
'	objExcel.Visible = False
'	Set objWorkbook = objExcel.Workbooks.Open(strTestSetPath)
'	
'	Set objWorksheet = objExcel.ActiveWorkbook.Worksheets("TestSet")
'	
'	'Get the max column occupied in the excel file 
'	intColCount = objWorksheet.UsedRange.Columns.Count
'	
'	'Read values
'	For i = 1 To intColCount
'		Select Case LCase(objWorksheet.Cells(1,i).Value)
'			Case "dcversion"
'				Environment("dcversion") = objWorksheet.Cells(2,i).Value
'			Case "trustversion"
'				Environment("buildnumber") = objWorksheet.Cells(2,i).Value
'			Case "agentversion"
'				Environment("agentversion") = objWorksheet.Cells(2,i).Value
'			Case "protocolversion"
'				Environment("protocolversion") = objWorksheet.Cells(2,i).Value
'			Case "serveruser"
'				Environment("serveruser") = objWorksheet.Cells(2,i).Value
'			Case "serverpassword"
'				Environment("serverpassword") = objWorksheet.Cells(2,i).Value
'			Case "dcip"
'				Environment("dcip") = objWorksheet.Cells(2,i).Value
'			Case "dcuser"
'				Environment("dcuser") = objWorksheet.Cells(2,i).Value
'			Case "dcpassword"
'				Environment("dcpassword") = objWorksheet.Cells(2,i).Value
'			Case "devicemodel"
'				Environment("devicemodel") = objWorksheet.Cells(2,i).Value
'			Case "deviceos"
'				Environment("platform") = objWorksheet.Cells(2,i).Value
'			Case "deviceosversion"
'				Environment("osversion") = objWorksheet.Cells(2,i).Value
'			Case "viewerorientation"
'				Environment("deviceorientation") = objWorksheet.Cells(2,i).Value
'			Case "viewerscale"
'				Environment("devicescale") = objWorksheet.Cells(2,i).Value
'			Case "appid"
'				Environment("appid") = objWorksheet.Cells(2,i).Value
'		End Select
'	Next
'	
'	'Close the Workbook
'	objExcel.ActiveWorkbook.Close
'	 
'	'Close Excel
'	objExcel.Application.Quit
'	 
'	Set objWorksheet = Nothing
'	Set objWorkbook = Nothing
'	Set objExcel = Nothing
'	Set objFSO = Nothing
'End Sub

'Print "Test"
'
'
'GetTestSet = ""
'strTestsFolderPath = "C:\Users\Naveen\Desktop\MobileLabsAutomationFramework\Tests\Trust\CrossPlatform\TrustBrowser\"
'		Set objFSO= CreateObject ("Scripting.FileSystemObject")
'		Set objParentFolder = objFSO.GetFolder(strTestsFolderPath)
'		Set objSubFolder = objParentFolder.SubFolders
'        For Each Folder in objSubFolder
'			GetTestSet = Trim(GetTestSet & strTestsFolderPath & Folder.Name & "||")
'		Next
'		
'		Set objFSO = Nothing
'		Set objParentFolder = Nothing
'		Set objSubFolder = Nothing
'		
'		GetTestSet = Left(GetTestSet,Len(GetTestSet)-2)
'		GetTestSet = Split(GetTestSet,"||")
'		Wait 2
