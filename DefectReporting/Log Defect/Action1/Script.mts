
'############ logging Bug for Failed Test Cases   #######################

SystemUtil.CloseProcessByName "iexplore.exe"
SystemUtil.Run "iexplore.exe","https://mobilelabs.atlassian.net/secure/Dashboard.jspa"
Wait 5

Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "{F11}"

Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").Link("Log In").Click
Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").Link("Log in directly").Click
Browser("System Dashboard - JIRA").Page("Log in - JIRA").WebEdit("os_username").Set "jyoti.handuja"
Browser("System Dashboard - JIRA").Page("Log in - JIRA").WebEdit("os_password").SetSecure "4faca303d1a96492922f39f36323f742b768ce587330"
Browser("System Dashboard - JIRA").Page("Log in - JIRA").WebButton("Log In").Click

'Set objQTP = CreateObject("QuickTest.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Finding current test  location 
'StrTestPath = objQTP.Test.Location
strTestPath = Environment("TestDir")
                                
'Do While  Replace(Split(StrTestPath,"\")(UBound(Split(StrTestPath,"\"))), " ", "") <> "DefectReporting"
'	StrTestPath = objFSO.GetParentFolderName(StrTestPath)
strFolderPath = objFSO.GetParentFolderName(strTestPath)
	
'		Exit Do

'Loop
'Dim excelPath                                                                      'the path to the excel file

'excelPath = StrTestPath                                             'where is the Excel file located? 

'For each excel file in strFolderPath, log a defect
	Set objFiles = ObjFSO.GetFolder(strFolderPath).Files
	For Each objFile  in objFiles
		If LCase(Right(objFile.Name, 4)) = "xlsx" OR LCase(Right(objFile.Name, 3)) = "xls" Then

			excelPath = objFile.Path
			Set objExcel = CreateObject("Excel.Application")
			objExcel.Visible = False
			Set objWorkbook = objExcel.Workbooks.Open(excelPath)
			
			Set objDict = CreateObject("Scripting.Dictionary")
			
			For Each objWorksheet in objExcel.ActiveWorkbook.Worksheets
			
				usedColumnsCount = objWorksheet.UsedRange.Columns.Count
						
				For iColCount = 1 To usedColumnsCount
					 objDict (CStr(objWorksheet.Cells(1,iColCount))) = CStr(objWorksheet.Cells(2,iColCount))
				Next
	
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").Link("Create new issue").Click                            
				Wait 5
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebEdit("WebEdit_4"). Click
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebEdit("WebEdit_4").Set "Mobile labs Trust"
				Wait 3
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebEdit("WebEdit_2").Click
				wait 2

				WshShell.SendKeys "Bug"
				WshShell.SendKeys "~"
				wait 7
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebEdit("summary").set objDict("Summary")
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebList("priority").Select "Major"
				Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebEdit("WebEdit_3").Set"QTP Add-In"		
				Browser("System Dashboard - JIRA").Page("Create Issue - JIRA").WebEdit("WebEdit").Set "Earl Adona"
				
				Browser("System Dashboard - JIRA").Page("Create Issue - JIRA").WebEdit("description").Set "Steps to Reproduce:" + vbLf + objDict("Details") + vbLf + "ExpectedResult:" + vbLf +objDict("ExpectedResult") + vbLf + "Actual Result:" + vbLf +objDict("ActualResult") + vbLf 
				'Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").WebButton("Create").Click
			 
			Next
			
			objWorkbook.Close
			objExcel.Quit
						   
			Set objWorkbook = Nothing
			 
			Set objExcel = Nothing                                                                 'done with the Excel object, release it from memory

		End If

	Next

	Set ObjFSO = Nothing
	Set objFiles = Nothing

Browser("System Dashboard - JIRA").Page("System Dashboard - JIRA").Sync
WshShell.SendKeys "{F11}"

'SystemUtil.CloseProcessByName "iexplore.exe"