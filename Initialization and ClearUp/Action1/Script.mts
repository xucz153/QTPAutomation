'####################################################################################################################
'  Name :     NGS Scripts Initialization for Starting Up 
' Objective           :     Initialize the Environments for running: 
	'1.Initialize Environment Parameters for running: Wait time, desired Browser, desired Environment URL
	'2.Close all desired browsers before running script and after Finished 
	'3. Open desired Browser in desired Environment: QA/Test
	'4.Maximize/Restore Browser Window
'Called by scripts:  Each script in Ecommerce
'Calling external actions:
	'Reusable_01_Public_Functions_WaitingForPageLoad [Reusable_01_Public_Functions_WaitingForPageLoad]
	'MaximizeBrowserWindow [Reusable_02_Public_Functions_BrowserWindowMaximizeControl]
	'RestoreBrowserWindow [Reusable_02_Public_Functions_BrowserWindowMaximizeControl]
' Author                 :     Loya.He
' Date Created   :     4/01/2014

'####################################################################################################################

'Load the target browsers and target url from external xml file
Environment.LoadFromFile("..\..\TestData\TestEnvironmentConfiguration.xml")


'Set the wait times as environment value
Environment.Value("sTime")= 10
Environment.Value("mTime")= 35
Environment.Value("lTime")= 60

'Set URLs as environment.value
Environment.Value("NGSQAEnv")=  "https://ngsq.rd.cat.com/SIS/servicecenter/#type=Home"
Environment.Value("NGSProdEnv")=  "https://ngs.cat.com/SIS/servicecenter/ "

userName = Environment.Value("userName")
passWord = Environment.Value("passWord")

'Define the browsers
Dim arrayTargetBrowser 
arrayTargetBrowser = array("iexplore.exe", "firefox.exe", "chrome.exe")

' Set the target Browser which the test scripts will be run against
Dim targetBrowserIndex 
targetBrowserIndex = Environment.Value("targetBrowserIndex")
Dim targetBrowser
targetBrowser = arrayTargetBrowser(targetBrowserIndex)

' Set the target URL which the test scripts will be run against
Dim targetEnvIndex
targetEnvIndex =  Environment.Value("targetEnvIndex")
Dim targetUrl 
targetUrl =  Environment.Value (targetEnvIndex)


'Set the targetBrowser as environment value to be used in other scripts
Environment.Value("targetBrowser") = targetBrowser
Environment.Value("targetEnv") = targetEnvIndex

'Close browsers before starting running
close_Browser(targetBrowserIndex)

'Open the speficied browser
If targetBrowser= "iexplore.exe" Then
	Environment.Value("objIE") =  createObject("InternetExplorer.application")
	' Open the IE browser and navigate to base test page set as targetURL
	open_Browser_Navigate_To_Base_Page Environment.value("objIE"), targetUrl
Else
	'Open browser FireFox or Chrome
	SystemUtil.Run targetBrowser, targetUrl
End If


'Verify if browser is displayed
isTargetBrowserDisplayed = Browser("BrowserCategoryLanding").Page("PageCategoryLanding").Exist
If  isTargetBrowserDisplayed = "True" Then
	Reporter.ReportEvent micDone, "Open browser", "Browser " &  targetBrowser & " is displayed."
Else 
	Reporter.ReportEvent micFail, "Open browser",  "Browser " &  targetBrowser & " is NOT displayed."
End If


'Wafting for page load completely

RunAction "Reusable_01_Public_Functions_WaitingForPageLoad [Reusable_01_Public_Functions_WaitingForPageLoad]", oneIteration

'Verify if category landing page is dispalyed.
'isTargetPageDisplayed = Browser("BrowserCategoryLanding").Page("PageCategoryLanding").WebElement("PageNameHome").Exist
'If isTargetPageDisplayed = "True" 	 Then
'	Reporter.ReportEvent micDone, "Navigate to test base page", "Page with TargetURL: " & targetUrl & " is displayed"
'Else 
'	Reporter.ReportEvent micFail, "Navigate to test base page", "Page with TargetURL: " & targetUrl & " is NOT displayed. "
'End If


'Maximize the Browser Window
RunAction "MaximizeBrowserWindow [Reusable_02_Public_Functions_BrowserWindowMaximizeControl]", oneIteration


'Login with desried user account
RunAction "Reusable_03_Public_Functions_Login [Reusable_03_Public_Functions_Login]", oneIteration