# MobileLabsAutomationFramework(UFT)

### How to get started with MobileLabs Automation Framework
		
This document describes what needs to be done in order to use the automation framework. Please follow the steps given below:

##Pre-requisites

1.	QTP 11.0 or later needs to be installed.
2.	A driver needs to be installed for PostgreSQL ANSI. Follow these steps:
	* Download the latest 32-bit driver from this link - https://www.postgresql.org/ftp/odbc/versions/msi/
	* Alternatively you can also install the driver included with the Framework > [Link](../Test/Tools/psqlodbc_x86.msi)	

### Device connectivity software

1.	Trust should be installed.
2.	Please make sure you install the right version of Trust downloaded from the deviceConnect server.

### Folder Structure

1.	Data – TBD, we don’t have any data driven scripts at the moment.
2.	Debug – Meant only to debug some code before being pushed to the final tests.
3.	DefectReporting – TBD, report defects to Jira automatically for failures. Not implemented as of now as it requires complex design and logic to determine which bugs need to go to Jira and also check if the bug/failure is already logged or not.
4.	Environment – File(s) that will be used to setup environment like what device model, os and app will the tests run on. TestLab.xlsx is the only file being used at the moment.
5.	FunctionLibraries – All qfls go in here and this folder will be added to QTP/UFT folder list at run-time.
6.	MasterScripts – Only ExecuteTestSet.vbs will be used for now which will read configuration from TestLab.xlsx and execute all the tests in the defined directory.
7.	Notification – TBD, will send email notifications on start and end of the execution.
8.	ORs – All object repositories go in here. This folder will be added to QTP/UFT folder list at run-time.
9.	Recovery – TDB, place for adding any recovery scenarios in future.
10.	Results – TBD, as at the moment the results are saved to a temp location.
11.	Tests – All tests will be kept in here. The folder structure is defined below:

	¬	Module – Like Trust run-time or Installer
	
	¬	Platform – Like iOS or Android
		
	¬	Application – Like PhoneLookup or Trust Browser
			
12.	Tools – All executables, dlls and other files needed at run-time will be placed here along with CLI files contained in a sub-folder. The files will be copied to %temp%\MobileLabsAutomation at run-time.
13.	TrustBuilds – TBD, for future when the Trust build will be installed automatically.

### Configuration

In the MobileLabsAutomationFramework folder go to Environment folder and open the TestLab.xlsx file.

1.	In the TestSet worksheet enter the following details.

-	dCVersion	
-	trustVersion
-	agentVersion
-	protocolVersion
-	serverUser
-	serverPassword
-	dCIP
-	dCUser
-	dCPassword
-	deviceModel
-	deviceOS
-	deviceOSVersion
-	testFolder – The folder where the tests are placed. To be selected from the dropdown list. Actual path will be created automatically at run-time. 
-	viewerOrientation
-	viewerScale
-	appID
-	nativeAutomation
-	addIns - Comma separated for more than one Add-in
2.	If any additional data is needed then it can be added in the Data worksheet.

### QTP/UFT setup
Make sure all absolute paths are removed, i.e. function libraries and object repositories should be referenced by the filename only and not by path. This will help in locating the files in the folders that ExecuteTestSet.vbs will add at run-time.

### Execution
Once the configuration has been done in TestLab.xlsx go to MasterScripts folder and run the ExecuteTestSet.vbs file.

### The Automation Framework will kick-off and do the following:
1.	Create an array of tests in the defined folder
2.	Run the tests one by one
3.	The rest of the configuration is read by ReadEnvironmentVariables function in GeneralFunctions.qfl
4.	The scripts would store tests results in %Temp\TrustTestResult
