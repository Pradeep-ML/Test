'Description : Open the testset excel file ,Delete all contents and then write the name of the test cases .
	'Declaring variables
	Dim strCurrentPath
	Dim TestcasePath
	'Dim CompletePath
	Dim StrPath
	Dim strCoulmnName
	Dim strRowName
	Dim ExcelColHead
	Dim TestFolderName
	Dim TestCaseName
	Dim strTestCasefoldercount
	Dim strcpath

	'Get current directory
	Set objWScript = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentPath = WScript.ScriptFullName
	 
	'Loop until "MobileLabs Automation Framework" folder is found
	blnParentFolderFound = True
	Do While Replace(Split(strCurrentPath,"\")(UBound(Split(strCurrentPath,"\"))), " ", "") <> "MobileLabsAutomationFramework"
		strCurrentPath = objFSO.GetParentFolderName(strCurrentPath)
		'Exit if reaches the system drive
		If InStr(1, strCurrentPath, "\") = 0 Then
			blnParentFolderFound = False
			Exit Do
		End If
	Loop
	 
	'Define the path of the Root Folder: <MobileLabs Automation Framework>
	If blnParentFolderFound Then
		If Right(strCurrentPath,1) <> "\" Then
			strCurrentPath = strCurrentPath & "\" '&"MobileLabs Tests"
		End If
		If Replace(objFSO.GetFolder(strCurrentPath).Name, " ", "") <> "MobileLabsAutomationFramework" Then
			WScript.Quit
		End If
	Else
		MsgBox "Error: CreateTestSet.vbs file is being executed from a wrong location: " & objWScript.CurrentDirectory
		WScript.Quit
	End If

	'Creating FSO object
   	TestcasePath = "MobileLabs Tests"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	CompletePath = strCurrentPath & TestcasePath
	
	StrPath = strCurrentPath & "Configuration\TestSet.xlsx"
'
'	'Code to clear the contents of the excel file 
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open (StrPath)
	
	For i=1 to objExcel.Worksheets.Count 
		Set objWorksheet = objExcel.Worksheets.Item(i)

		If  objExcel.Worksheets.Item(i).name <> "Data"Then
			strCoulmnName = ObjWorkSheet.USedRange.Columns.Count
			strRowName = ObjWorkSheet.USedRange.Rows.Count
			'Code to get header name of column based on user range for column
			If strCoulmnName > 26 Then
				T = (strCoulmnName Mod 26)
				strCoulmnName = (strCoulmnName - T) / 26
				ExcelColHead = Chr(64 + strCoulmnName) & Chr(64 + T)
				Else
				ExcelColHead = Chr(64 + strCoulmnName)
			End If
			coulmnToDelete = ExcelColHead & strRowName
			If  strRowName = 1 Then
				coulmnToDelete = ExcelColHead & (strRowName +1)
				'Delete the contents of the excel file
				objWorksheet.Range("A2:" & coulmnToDelete).Delete
				Else
				'Delete the contents of the excel file
				objWorksheet.Range("A2:" & coulmnToDelete).Delete
			End If
		End If
	Next

	objExcel.ActiveWorkbook.Application.DisplayAlerts = false
	objExcel.ActiveWorkbook.Save
	objWorkbook.Close
	objExcel.Quit
	
	Set objWorksheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = nothing

	''''''''''''''''''''''''''''''''''''''''''''''
	'Checking validity of path entered
	If  objFSO.FolderExists(CompletePath) Then
		Set ParentFolders1 = objFSO.GetFolder(CompletePath)
		Set SubFolder = ParentFolders1.SubFolders
		i = 0
		For Each FolderName in SubFolder
			'Save the name of the folder in variable
			TestFolderName = FolderName.name
			'Call get foldername function to name of the subfolders
			Call GetFolderName
			i = i+1
		Next        
	End If


	'##########################################################################################################################
	''@Function:        GetFolderName
	''@Description: 	Get the  folder name
	''@Created By: 		Saurabh Ahuja
	''@Created On: 		06/08/2012
	''--------------------------------------------------------------------------------------------------------------------------
	''@Param Name: 		TestCaseName
	''@Param Type: 		String
	''@Param Drtn: 		out
	''@Param Desc: 		Name of  the Test case to be write to excel
	''--------------------------------------------------------------------------------------------------------------------------
	''@Param Name: 		strTestCasefoldercount
	''@Param Type: 		String
	''@Param Drtn: 		out
	''@Param Desc: 		Expected behaviour of the method which failed
	''--------------------------------------------------------------------------------------------------------------------------
	''@Param Name: 		Devicename
	''@Param Type: 		String
	''@Param Drtn: 		In
	''@Param Desc: 		Device name to whom test case belongs
	''--------------------------------------------------------------------------------------------------------------------------

	Function GetFolderName
		FolderPathName = CompletePath & "\"& TestFolderName
		MsgBox FolderPathName
		strcpath = FolderPathName & "\"
		Set objFSO2= CreateObject ("Scripting.FileSystemObject")
		Set ParentFolders1 = objFSO2.GetFolder(FolderPathName)
		Set SubFolder1 = ParentFolders1.SubFolders
        For each foldername1 in SubFolder1
			
			Devicename = FolderName1.name
			FolderPathName1 = FolderPathName & "\" & Devicename
			Set ParentFolders1 = objFSO2.GetFolder(FolderPathName1)
			Set SubFolder2 = ParentFolders1.SubFolders
			strTestCasefoldercount = SubFolder2.Count
				l = 0
				'Check whether there is any sub folder or not
				If  SubFolder2.Count > 0Then
					'SubCount = SubFolder2.Count
					l = 2
					r = 1
				
					'Check for the name of the sub folder
					For Each FolderName2 in SubFolder2
						TestCaseName = FolderName2.name
						'Call excel function to write the data
						call WriteExcelData(TestCaseName,strTestCasefoldercount,l,r,strcpath,Devicename)
						l = l+1
						r = r+1
					Next  
				End If
		Next

	End Function
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
	Sub WriteExcelData(TestCaseName,strTestCasefoldercount,l,r,strcpath,Devicename)
		'Check for the file existence

		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		Set objFSO3= CreateObject ("Scripting.FileSystemObject")
		If objFSO3.FileExists(StrPath) <> true  then
			set objWorkbook = objExcel.Workbooks.Add
			objExcel.ActiveWorkbook.SaveAs (StrPath)
			'Delete the additional sheet if any
			For sheetCounter = 1  to objExcel.Worksheets.Count-1
				objExcel.Worksheets(sheetCounter).Delete
			Next
			Set objWorksheet = objExcel.Sheets.Item(1)
			objWorksheet.Name = TestFolderName		
			
		Else
			objExcel.Workbooks.Open StrPath
		End if


		blnSheetExist = False
		'if name of the sheet not present on then add sheet with name of the folder
		For i=1 to objExcel.Worksheets.Count
			If objExcel.Worksheets(i).Name =  TestFolderName Then
				blnSheetExist = True
				Exit For
			End If
		Next

		If blnSheetExist = False Then
			Set objWorksheet = objExcel.Worksheets.Add
			objWorksheet.Name = TestFolderName
		End If


		'Code to create list validation at run rime
		Const xlValidateList = 3  
		Const xlThin = 2  
		Const xlContinuous = 1  
		Set objSheet = objExcel.Worksheets(TestFolderName) 
		objSheet.Activate 
		strRowUsedRange = objSheet.USedRange.Rows.Count
		objExcel.Cells(strRowUsedRange+1, 3).Validation.Add xlValidateList,,,"YES,NO"  

		'Write test case name on excel sheet
		Set objWorksheet = objExcel.Worksheets.Item(TestFolderName)
		strRowUsedRange = ObjWorkSheet.USedRange.Rows.Count
		objWorksheet.Cells(strRowUsedRange+1,1).Value = strcpath & Devicename & "\" & TestCaseName
		objWorksheet.Cells(strRowUsedRange+1,3).Value = "YES"
		objWorksheet.Cells(strRowUsedRange+1,2).Value = Devicename

		Set objRange = objWorksheet.UsedRange
		objRange.EntireColumn.Autofit()
		
		objExcel.ActiveWorkbook.Save
		
		objExcel.Quit
		Set objExcel = nothing
	End Sub


	Set objFSO = Nothing
	Set objFSO2 = Nothing
	Set objFSO3 = Nothing

	MsgBox "Completed..!!"
































