'#######################################Send Notification###################################
'Description: 	This script sends a notification to the specified target users about the 
'				Automation Framework being kicked off.
'###########################################################################################

'Declare all variables
Dim strCurrentPath
Dim blnParentFolderFound
Dim strCleanupPath
Dim strEnvironmentPath
Dim strBuild
Dim strPlatforms
Dim strTestSetPath
Dim strSystemInfo

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
		strCurrentPath = strCurrentPath & "\"
	End If
	If Replace(objFSO.GetFolder(strCurrentPath).Name, " ", "") <> "MobileLabsAutomationFramework" Then
		MsgBox "MobileLabs Automation Framework folder was not found!"
		WScript.Quit
	End If
Else
	MsgBox "Error: SendNotification.vbs file is being exeecuted from a wrong location: " & objWScript.CurrentDirectory
	WScript.Quit
End If

'Get data from Environment.xlsx
strEnvironmentPath = strCurrentPath & "\Environment\EnvironmentVariables.xlsx"
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open(strEnvironmentPath)
Set objWorksheet = objExcel.ActiveWorkbook.Worksheets("Variables")

For i = 1 to objWorksheet.UsedRange.Rows.Count
	Select Case UCase(objWorksheet.Cells(i,1))
		Case "LATEST_BUILD"
			strBuild = objWorksheet.Cells(i,2)
	End Select
Next

objWorkbook.Close

'Get name of platforms on which execution is about to start: Get the names of the worksheets in TestSet.xlsx
strTestSetPath = strCurrentPath & "\Configuration\TestSet.xlsx"
Set objWorkbook = objExcel.Workbooks.Open(strTestSetPath)

For Each objWorksheet In objExcel.ActiveWorkbook.Worksheets
	strPlatforms = strPlatforms & objWorksheet.Name & ", " 
Next

strPlatforms = Left(strPlatforms, Len(strPlatforms) - 2)

objWorkbook.Close

objExcel.Quit

Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing

'Get System Info
Set objWMIService = GetObject( "winmgmts:\\.\root\cimv2" )
Set colItems = objWMIService.ExecQuery( "Select * from Win32_ComputerSystem")', , 48 )

For Each objItem in colItems
   For Each objMethod In objItem.Properties_
		Select Case UCase(objMethod.Name)
			Case "CURRENTTIMEZONE"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "DESCRIPTION"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "DNSHOSTNAME"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "Domain"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "MANUFACTURER"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "NAME"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "NUMBEROFPROCESSORS"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "PARTOFDOMAIN"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "PRIMARYOWNERNAME"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "SYSTEMTYPE"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & objMethod.Value & VBNewLine

			Case "TOTALPHYSICALMEMORY"
				strSystemInfo = strSystemInfo & objMethod.Name & ": " & CInt(objMethod.Value/1073741824) & " GB" & VBNewLine

		End Select
   Next
Next

'Send notification
Const fromEmail = "mobilelabsQA@gmail.com"
Const password = "QA@D-25Noida"

Set objEmail = CreateObject("CDO.Message")
objEmail.From = fromEmail
objEmail.To = "naveen.chauhan@pyramidconsultinginc.com"
objEmail.Subject = "Automation Framework has been kicked off on: " & strBuild
objEmail.TextBody = "Test execution has started:" & VBNewLine & "Time: " & Time & VBNewLine & "Platforms: " &_
 strPlatforms & VBNewLine & "Build: " & strBuild & VBNewLine & VBNewLine & "Given below is the machine info: " &_
 VBNewLine & strSystemInfo

If WScript.Arguments.Count > 3 Then
objEmail.AddAttachment WScript.Arguments.Item(3)
End If

Set objConfigEmail = objEmail.Configuration
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = fromEmail
objConfigEmail.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
objConfigEmail.Fields.Update

objEmail.Send

Set objEmail = nothing
Set objConfigEmail = nothing