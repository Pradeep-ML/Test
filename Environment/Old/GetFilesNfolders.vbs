On Error Resume Next 
strFolderPath = InputBox("Enter Folder Path")

If strFolderPath <> "" Then
	' Create Excel object
	Set objExcel = CreateObject("Excel.Application")
	
	''Disable Alerts
	objExcel.DisplayAlerts = False
	Excel.Application.EnableEvents = False 
	objExcel.Visible = False
	
	'Define sheet name
	arrSheetName = Split(strFolderPath, "\")	
	strSheetName = Replace(arrSheetName(UBound(arrSheetName)-1) & "_" & arrSheetName(UBound(arrSheetName)), " ", "")
		
	''Check if file exists, Then delete file
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strCurrentPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
	Set objFolder = objFSO.GetFolder(strCurrentPath)
	strFilePath = strCurrentPath & "\InstallerDirs.xlsx"
	
	''Create a Workbook object
	If Not(objFSO.FileExists(strFilePath)) Then
		Set objWorkbook = objExcel.Workbooks.Add()
	Else
		Set objWorkbook = objExcel.Workbooks.Open(strFilePath)
	End If
	
	blnSheetFound = False
		
	'Clear any existing sheet that matches the strSheetName
	For Each objWorksheet in objWorkbook.Worksheets
		If InStr(1, objWorksheet.Name, strSheetName, 1) > 0  Then		
			objWorksheet.Cells.Clear
			blnSheetFound = True
			Exit For
		End If
	Next
	
	If Not(blnSheetFound) Then
		Set objWorksheet = objWorkbook.Worksheets.Add()
		objWorksheet.Name = strSheetName
	End If
	
	'Set Header Values
	Set objWorksheet = objWorkbook.Worksheets(strSheetName)
	
	objWorksheet.cells(1,1) = "FileORFolder"
	objWorksheet.cells(1,2) = "Type"
	objWorksheet.cells(1,1).Font.Bold = True
	objWorksheet.cells(1,2).Font.Bold = True
	
	WriteInExcel strFolderPath
	
	'Save the changes
	objWorkBook.SaveAs(strFilePath)         
	objWorkBook.Close
	objExcel.Quit
	
	'Clear all the references to the objects
	Set objWorkBook = Nothing        
	Set objExcel = Nothing 
	Set objWorksheet = Nothing

Public Sub WriteInExcel(strFolderPath)

    'Get Rows Count and wirte at the end
	RowCount=objWorksheet.UsedRange.rows.count + 1

	Set Foldr = objFSO.GetFolder(strFolderPath)
    Set subFoldr = Foldr.SubFolders
	Set Fils = Foldr.Files
	'' Write values in Excel
	For Each f1 in subFoldr
		objWorksheet.cells(RowCount ,1) = f1.Path
		objWorksheet.cells(RowCount ,2) = "Folder"
		RowCount = RowCount + 1
	Next
	
	For Each f2 in Fils
		objWorksheet.cells(RowCount  ,1) = f2.Path
		objWorksheet.cells(RowCount ,2) = "File"
		RowCount = RowCount + 1
	Next
	
	if subFoldr.Count > 0 Then
		For Each sf in subFoldr
			Call WriteInExcel(strFolderPath & "\" & sf.Name)
		Next 
	End If
	 
End Sub

msgBox "File Created Successfully !!"

Else
	msgBox "Please Enter a valid Path"
End If