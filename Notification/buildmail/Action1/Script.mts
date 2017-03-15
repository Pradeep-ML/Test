strCurrentPath = Environment("TestDir")
Dim CurrentPath
CurrentPath = Split(strCurrentPath,"\")
UppBound = Ubound(CurrentPath)
For i = 0 to UppBound-2
	temp = CurrentPath(i)
	strCurrentPaths = strCurrentPaths & temp &"\"
Next

StrExcelPath = strCurrentPaths & "\Environment\EnvironmentVariables.xlsx"
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.Workbooks.Open StrExcelPath
'Pointing to Variables excel
set ObjWorkSheet = objExcel.Worksheets("Variables")
objRange = ObjWorkSheet.USedRange.Rows.Count
objRangeColumn = ObjWorkSheet.USedRange.Columns.Count
For i = 1 to objRange
	ColumnName = objWorksheet.Cells(objRange,1).value
	If  ColumnName = "LATEST_BUILD" Then
		Buildnumber = objWorksheet.Cells(objRange,2).value
    	End If
	Exit for
Next

'Save the excel file
objExcel.ActiveWorkbook.Save
'Close the excel file
objExcel.Quit
Set objExcel = nothing

'Close the all opened browser
SystemUtil.CloseProcessByName "iexplore.exe"
SystemUtil.Run "iexplore.exe","www.gmail.com"
Wait 5

Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys "{F11}"

'Description created for Mail browser
Set BrowserMail=Description.Create()
BrowserMail("micclass").Value="Browser"
BrowserMail("number of tabs").Value="1"

'Description created for Mail Page 
Set PageMail=Description.Create()
PageMail("micclass").Value="Page"
PageMail("title").value="Gmail: Email from Google"

 'Description created for Email id
Set EditEmail=Description.Create()
EditEmail("micclass").Value="WebEdit"
EditEmail("name").Value="Email"

 'Description created for Email Pass
Set EditPass=Description.Create()
EditPass("micclass").Value="WebEdit"
EditPass("name").Value="Passwd"

 'Description created for Sign in button
Set ButsignIn=Description.Create()
ButsignIn("micclass").Value="WebButton"
ButsignIn("name").Value="Sign In"

 'Set the email id
UsrEdit=Browser(BrowserMail).Page(PageMail).WebEdit(EditEmail).Set("mobilelabsqa@gmail.com")
wait 3
'set password
PwdEdit=Browser(BrowserMail).Page(PageMail).WebEdit(EditPass).SetSecure( "4fc2fb884dedcd126c7e8abc315df92ed3a2d86f2bcfca64d41e07c90719")
wait 2
'Click on Sign in button
SignIn=Browser(BrowserMail).Page(PageMail).WebButton(ButsignIn).Click
wait 7

 'Description created for Inbox
Set PageGmail=Description.Create()
PageGmail("micclass").Value="Page"
PageGmail("title").value=".*@gmail.com - Gmail"
 
 
Setting.WebPackage("ReplayType") = 2
    Browser("title:=Gmail.*").WebElement("innertext:=COMPOSE", "index:=1").Click
Setting.WebPackage("ReplayType") = 1
 
 'Description created for Mail ID text area
Set EditToMail=Description.Create()
EditToMail("micclass").Value="WebEdit"
EditToMail("html tag").Value="TEXTAREA"
EditToMail("name").Value="to"

 'Enter email id
ToEdit=Browser(BrowserMail).Page(PageGmail).WebEdit(EditToMail).Set ( "naveen.chauhan@pyramidconsultinginc.com")', manisha.miglani@pyramidconsultinginc.com, lavesh.verma@pyramidconsultinginc.com, saurabh@pyramidconsultinginc.com, jyoti.handuja@pyramidconsultinginc.com, Nitin.Mittal@pyramidconsultinginc.com")
'ToEdit=Browser(BrowserMail).Page(PageGmail).WebEdit(EditToMail).Set ( "saurabh@pyramidconsultinginc.com")
 'Description created for  Subject filed
Set EditMail=Description.Create()
EditMail("micclass").Value="WebEdit"
EditMail("html tag").Value="INPUT"
EditMail("name").Value="subject"

 'Enter value of  subject
ToEdit=Browser(BrowserMail).Page(PageGmail).WebEdit(EditMail).Set("Automation Framework kicked off at " & Time & " on " & Buildnumber)

wait 5
 'Description created for Mail Text Area
Set WebEleBody=Description.Create()
WebEleBody("micclass").Value="WebElement"
WebEleBody("html tag").Value="BODY"
WebEleBody("class").Value="editable LW-avf"

 'Click on Text Body
Browser(BrowserMail).Page(PageGmail).WebElement(WebEleBody).Click
 
wait 5
With Browser("title:=Gmail.*")
'    .WebElement(WebEleBody).Object.innerText = "Download Latest Build : " & Buildnumber &" From Location: " &"\\10.4.4.2\MobileLabsQA\MobileLabs AutomationFramework\TrustBuilds"
	.WebElement(WebEleBody).Object.innerText = "Test execution started for " & Buildnumber & " on machine: " & Environment("LocalHostName")
End With

wait 5
 'Click on send

 Setting.WebPackage("ReplayType") = 2
    Browser("title:=Gmail.*").WebElement("innertext:=Send", "index:=1").Click
Setting.WebPackage("ReplayType") = 1
 
wait 5

 'Click on accounts
Set LinksignoutMail=Description.Create()
LinksignoutMail("micclass").Value="Link"
LinksignoutMail("text").Value="mobilelabsqa@gmail.com"

SignO=Browser(BrowserMail).Page(PageGmail).Link(LinksignoutMail).Click
wait 5
  'Click on sign out
Browser("title:=Gmail.*").Page("micclass:=Page").Link("innertext:=Sign Out", "class:=gbqfbb").Click
wait 5
Browser("title:=Gmail.*").Page("micclass:=Page").Sync
WshShell.SendKeys "{F11}"
SystemUtil.CloseProcessByName "iexplore.exe"
 
 


























