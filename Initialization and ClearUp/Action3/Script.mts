'#################################################################################################################
'  Name :     Initialization and ClearUp
' Objective           :     Get the targetBrowser and  test environment from xml file, set as environment values.
' Author                 :     Sarah Shi
' Date Created   :     1/23/2013
' Date Updated :     02/01/2013
'Updated by       :     Loya He
'#################################################################################################################

'Load the target browsers and target url from external xml file
Environment.LoadFromFile("..\..\TestData\TestEnvironmentConfiguration.xml")

'Set the wait times as environment value
Environment.Value("sTime")= 10
Environment.Value("mTime")= 35
Environment.Value("lTime")= 60

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

'Set the targetBrowser as environment value to be used in other scripts
Environment.Value("targetBrowser") = targetBrowser
Environment.Value("targetEnv") = targetEnvIndex

'Close browsers before starting running
close_Browser(targetBrowser)
