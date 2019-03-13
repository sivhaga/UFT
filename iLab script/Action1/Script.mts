'Initializing the variables 
Call IntializeVariables()  


'In case of any error occured in between the steps, move on to next step
On Error Resume Next

'Declaration of variable
Public strResult

'Getting Time value for Reporting Functions - Start time of Test
StartTime = Now 
Environment.Value("ENV_EXCEUTION_START_TIME") = StartTime

'Formatting the Date and Time Value
gCurrDate  = F_FormatDateTime(StartTime,"HeaderDate")
gCurrTime  = F_FormatDateTime(StartTime,"HeaderTime")

'Giving the Application name that will be displayed in Consolidated Report
Environment.Value("Application")	= 	"iLab"

'To write a consolidated Report for our Test Run - This will be helpful for Batch Run
Call F_ConsolidatedReportFileHeader(gCurrDate,gCurrTime)

'Extracting details from the DRIVER file 
strModuleName = "DRIVER_iLab"
strDataFilePath = Environment.Value("ENV_DATA_FOLDER_PATH") & strModuleName

'Considering Driver Sheet as Database to retrive the corresponding values by using SQLQuery
strConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DBQ=" & strDataFilePath &";ReadOnly =True;"
strSQLQuery = "Select * from [EXECUTION$] where TC_EXECUTE='Yes' or TC_EXECUTE='yes' or TC_EXECUTE='yes'"

Set st_DriverDataFields = getRecordSet(strConnectionString,strSQLQuery)

'Initiallise the Variables
Environment.Value("TC_TOTAL")  		= 0 
Environment.Value("TC_PASSED") 		= 0 
Environment.Value("TC_FAILED") 		= 0
Environment.Value("Step_Number")	= 0

'Start Performing the Action on the Selected scenarios
Do Until st_DriverDataFields.EOF
'Extracting the Values from EXECUTION sheet in DRIVER fle   		     
	Environment.Value("TC_EXECUTE_Flag")	=	UCase(Trim(st_DriverDataFields.Fields.Item("TC_EXECUTE")))
	Environment.Value("TC_NAME")			=	UCase(Trim(st_DriverDataFields.Fields.Item("TESTCASE")))
	Environment.Value("TC_DESC")			=  	UCase(Trim(st_DriverDataFields.Fields.Item("DESCRIPTION")))
	Environment.Value("ScenarioNumber")		=	UCase(Trim(st_DriverDataFields.Fields.Item("TESTCASE")))
   	Environment.Value("TEST_CONFIG_ID")		= 	Trim(st_DriverDataFields.Fields.Item("TC_ID"))
	Environment.Value("TC_TOTAL") 			= 	Environment.Value("TC_TOTAL") + 1
	
'Creating FileSystemObject for creating Report Folder
	Set fso = CreateObject("Scripting.FileSystemObject")

'Creating the Folder, if it doesn't exist
	If NOT fso.FolderExists(Environment.Value("ENV_REPORT_FOLDER_PATH_DATE")) Then
		fso.CreateFolder(Environment.Value("ENV_REPORT_FOLDER_PATH_DATE"))
	End If

'Creating Report folder on TestCase Name
	Environment.Value("ENV_REPORT_FOLDER_PATH_TC_NAME")	= Environment.Value("ENV_REPORT_FOLDER_PATH_DATE")&"\"&Environment.Value("TC_NAME")
	
'Creating the Folder on TestCase Name, if it doesn't exist
	If NOT fso.FolderExists(Environment.Value("ENV_REPORT_FOLDER_PATH_TC_NAME")) Then
		fso.CreateFolder(Environment.Value("ENV_REPORT_FOLDER_PATH_TC_NAME"))
	Else
		fso.DeleteFolder(Environment.Value("ENV_REPORT_FOLDER_PATH_TC_NAME"))
		fso.CreateFolder(Environment.Value("ENV_REPORT_FOLDER_PATH_TC_NAME"))
	End If

'Making the Folder to open
	Environment.Value("ENV_REPORT_FOLDER_PATH")=Environment.Value("ENV_REPORT_FOLDER_PATH_TC_NAME")&"\"

' Closing the reference object fso
 	Set fso = Nothing

'Making he Report folder path as Screenshot folder path	
	Environment.Value("ENV_SCREENSHOT_FOLDER_PATH") = Environment.Value("ENV_REPORT_FOLDER_PATH")
	
'HTML Reporting file for Individual Test Case
	Call F_ReportFileHeader(gCurrDate,gCurrTime)
	Call F_ScenarioHeader(Environment.Value("TC_NAME"),Environment.Value("TC_DESC"))

'Initialize the Execute command result as True
    strResult = "True"
 
'Initialising the Passed and Failed Test Steps count 
    Environment.Value("No_Fail_Steps") = 0
    Environment.Value("No_Pass_Steps") = 0
    
'Getting Test Case Starting Time    
	Environment.Value("ENV_EXCEUTION_TC_START_TIME") = Now
	
'Execute the each function in sequence to complete the test cases			
	For intCurrentColumn = 4 to st_DriverDataFields.Fields.Count - 1 ' Function Name starts at Column 5.
		
		strFunction_Name = st_DriverDataFields.Fields.Item(intCurrentColumn)
	
'Verify the Function name is available or Null
		If  Trim(strFunction_Name)  <> ""  Then

'Executing the function taken from the 'EXECUTION' sheet of DRIVER file  for corresponding test case
			Execute "strResult = " & strFunction_Name & "()"
	
'Validating the Result of Function Execution and determining the TC_RESULT status
			If strResult <> True Then       
				Environment.Value("TC_RESULT") = "FAIL"
'Stops the test case execution if any function has been failed
				Exit For
			Else
				Environment.Value("TC_RESULT") = "PASS"
			End If  
		End If
	Next
	
'Get Execution End Time for a Test case
	Environment.Value("ENV_EXCEUTION_TC_END_TIME")=Now
	
    If Environment.Value("TC_RESULT") = "FAIL" Then
'HTML Reporting Footer - For Individual Report
		Call F_ScenarioFooter(Environment.Value("TC_NAME"),"FAIL")
'HTML Reporting Footer - For Consolidated
		Call F_ConsoldatedScenarioFooter(Environment.Value("TC_NAME"),"FAIL")
'Overall Failed count calculation
		Environment.Value("TC_FAILED") = Environment.Value("TC_FAILED") + 1
'Else part for Passed Test cases
	Else
		Call F_ScenarioFooter(Environment.Value("TC_NAME"),"PASS")
		Call F_ConsoldatedScenarioFooter(Environment.Value("TC_NAME"),"PASS")
		Environment.Value("TC_PASSED") = Environment.Value("TC_PASSED") + 1
	End If  

'Moving to the Next Record
	st_DriverDataFields.MoveNext
	
Loop 

'Check for TC_EXECUTE Flag in Driver. If None of the TC_EXECUTE Flag was set to Yes then Report
If Environment.Value("TC_TOTAL") <> 0 Then 
	EndTime = Now 
	Environment.Value("ENV_EXCEUTION_END_TIME") = EndTime
	Call  F_ConsolidatedExecutionReport(Environment.Value("TC_TOTAL") ,Environment.Value("TC_PASSED"),Environment.Value("TC_FAILED") )
Else
	Print "None of the 'TC_EXECUTE' Flag was set to Yes"
End If
	
'*************************************************************************************************************
' 	Function    : IntializeVariables
' 	Purpose     : Intialize all Environment Variables
'	Inputs      : Folder path in the 'PathConfiguration' sheet in "DRIVER_iLab" file
' 	***************************************************************************************************************
Public Function IntializeVariables()

'Assigning the sheet name as action name
	strDataSheetName=Environment.Value ("ActionName")
'Import the local sheet
	Datatable.AddSheet strDataSheetName
	Datatable.ImportSheet  "C:\UFT\Data\DRIVER_iLab.xls","Path_Configuration", strDataSheetName

'Fetching the Row count 
	intScenarioCount = Datatable.GetSheet(strDataSheetName).GetRowCount

'Flag with null value
	strScenarioFlag=""

'For Setting the Row for the Current Scenario
	For intScenarioIteration=1  to intScenarioCount
		Datatable.GetSheet(strDataSheetName).setCurrentRow(intScenarioIteration)
		If Trim(DataTable.Value("File_Control",strDataSheetName))="Yes" Then
			strScenarioFlag="1"
			Exit For
		End If
	Next

'Current Date for Reporting file
	gCurrDate  = F_FormatDateTime(Now,"HeaderDate")
	
'For Handling the absence of Channel information
	If strScenarioFlag <> "1" Then
		Msgbox  "Folder Path is not defined in config file"
	End If

	Environment.Value("ENV_DATA_FOLDER_PATH")					= 	Trim(DataTable.Value("ENV_DATA_FOLDER_PATH",strDataSheetName))
	Environment.Value("ENV_REPORT_FOLDER_PATH_")				= 	Trim(DataTable.Value("ENV_REPORT_FOLDER_PATH",strDataSheetName))
	Environment.Value("ENV_CONSOLIDATED_REPORT_FOLDER_PATH_")	= 	Trim(DataTable.Value("ENV_CONSOLIDATED_REPORT_FOLDER_PATH",strDataSheetName))
	Environment.Value("ENV_CONFIG_FILE_PATH")					=  	Trim(DataTable.Value("ENV_CONFIG_FILE_PATH",strDataSheetName))
	Environment.Value("ENV_OR_FOLDER_PATH")						=  	Trim(DataTable.Value("ENV_OR_FOLDER_PATH",strDataSheetName))
	Environment.Value("ENV_LIBRARY_FOLDER_PATH")				=  	Trim(DataTable.Value("ENV_LIBRARY_FOLDER_PATH",strDataSheetName))
	Environment.Value("ENV_FOLDER_PATH")						=  	Trim(DataTable.Value("ENV_FOLDER_PATH",strDataSheetName))
	Environment.Value("ENV_RESOURCES_FOLDER_PATH") 				=   Trim(DataTable.Value("ENV_RESOURCES_FOLDER_PATH",strDataSheetName))

'Making Report Folder Name ending with Current Date
    Environment.Value("ENV_REPORT_FOLDER_PATH_DATE")			=	Environment.Value("ENV_REPORT_FOLDER_PATH_")&"_"&gCurrDate
    Environment.Value("ENV_CONSOLIDATED_REPORT_FOLDER_PATH")	= 	Environment.Value("ENV_CONSOLIDATED_REPORT_FOLDER_PATH_")&"_"&gCurrDate&"\" 'To keep the Consolidated Report under the Report folder that ends with Current Date
    
End Function	



