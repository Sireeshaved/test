'*************************' Demo Test Case for Metcash '*************************
'*************************' Designed By Koventhan '*************************
'the data will be used from UFT data Tables. Report also will be generated same.
'*************************' Pre Requisites :'*************************
'	1. All the Values will be taken from the input data sheet. Make Sure you give all the login details properly
'	2. Input data needs to be given in Data sheet. Make sure proper values are passed.'Launch the Chrome Browser
' Load the Dynamics 365 URL
On Error Resume Next
Iter = Environment.Value("TestIteration")-2
'msgbox(DataTable.Value("D365_URL","Sheet1"))
SystemUtil.Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" ,"https://almnz-uat.sandbox.operations.dynamics.com/"
Set Homepage =  Browser("Dashboard -- Finance and").Page("Dashboard -- Finance and")
'wait for page to get loaded to perform next operation
wait 120
Browser("Dashboard -- Finance and").Sync
Call TakeScreenShot()
'Click On Modules and wait for element to get loaded
Browser("Dashboard -- Finance and").Page("Dashboard -- Finance and").WebElement("Modules").Click
wait 2
'click on Accounts Receivable and wait for Page to Load
Browser("Dashboard -- Finance and").Page("Dashboard -- Finance and").WebElement("Accounts receivable").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebButton("Collapse all").Click
wait 2
'click on Orders from Accounts Receivable Menu and wait for Page to Load
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Orders").Click
wait 2
'CLick on All Sales Order if Order is collapsed or Expanded
If Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("All sales orders").Exist(2) Then
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("All sales orders").Click
Else
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Orders").Click
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("All sales orders").Click
End If
wait 3
Call TakeScreenShot()
'Click on the Web Portal Number and enter the value as Order Number from Legacy Portal and Click on Apply Button
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Web Portal Number").Click
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Sync
wait 3
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebEdit("Filter field: Web Portal").Set DataTable.Value("LegacyOrderNumber","ResultSheet")
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Sync
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("FIlter_Apply_Button").Click
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Sync
wait 3
Call TakeScreenShot()
'Click on the Sales Order Number
'SalesOrd=Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Link("SalesOrderNumber").GetROProperty("acc_name")
'SalesDetails = Split(SalesOrd," ")
'SalesOrderNumber = Trim(SalesDetails(2))
SalesOrd=Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Link("SalesOrderNumber").GetROProperty("title")
SalesDetails = Mid(SalesOrd,1,10)
SalesOrderNumber = Trim(SalesDetails)
	If SalesOrderNumber <> "" Then
			Call AddResults(Iter,"D365_SalesOrderNumber",SalesOrderNumber,"Sales Order Number Generated")
			wait 2
	Else
			Call AddResults(Iter,"D365_SalesOrderNumber",SalesOrderNumber,"Sales Order Number Not Generated")
			Environment.Value("FailureReason") = Environment.Value("FailureReason")  & "Sales Order Number Not Generated"
			Environment.Value("FailedApplication") = Environment.Value("FailedApplication") & "D365 Portal"
	End If
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Link("SalesOrderNumber").Click
	Browser("All sales orders -- Finance").Sync
	wait 4
'Click on Sales Order Header and Validate Values
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("SalesOrderHeader").Click
wait 2
Browser("All sales orders -- Finance").Sync
'Order Validated and Order Confirmed Values Validation
If Trim(Ucase(Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("OrderValidatedValue").GetROProperty("innertext"))) = "YES" Then
	Call AddResults(Iter,"D365_Order Validated Value","Yes","Pass")
Else
	Call AddResults(Iter,"D365_Order Validated Value","No","Fail")
	Environment.Value("FailureReason") = Environment.Value("FailureReason")  & "Order Validated Value is NO"
	Environment.Value("FailedApplication") = Environment.Value("FailedApplication") & "D365 Portal"
End  If

If Trim(Ucase(Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("OrderCOnfirmedValue").GetROProperty("innertext"))) = "YES" Then
	Call AddResults(Iter,"D365_Order Confirmed Value","Yes","Pass")
Else
	Call AddResults(Iter,"D365_Order Confirmed Value","No","Fail")
	Environment.Value("FailureReason") = Environment.Value("FailureReason")  & "Order Confirmed Value is NO"
	Environment.Value("FailedApplication") = Environment.Value("FailedApplication") & "D365 Portal"
End  If
Call TakeScreenShot()
'27527346   0000252785
'Navigate to modules-warehousemanagement-inquiryandreport-ALM-salesOrderSummary
'Click On Modules and wait for element to get loaded
Browser("Dashboard -- Finance and").Page("Dashboard -- Finance and").WebElement("Modules").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Warehouse management").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebButton("Collapse all").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Inquiries and reports").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("ALM").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Sales order summary").Click
wait 15
Call TakeScreenShot()

'Filter with sales Order Number
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Sales order Filter").Click
wait 3
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebEdit("Filter field: Sales orderNumber").Set "0000"&DataTable.Value("D365_SalesOrderNumber","ResultSheet")
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebButton("Sales Order FIlter Apply").Click
wait 9
'Capture WareHouse Order Number
WareHouse = Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Link("WarehouseOrderNumber").GetROProperty("acc_name")'
WareHouseOrder = Split(WareHouse," ")
WareHouseOrderNumber = Trim(WareHouseOrder(2))
If WareHouseOrderNumber <> "" Then
	Call AddResults(Iter,"D365_WareHouseOrderNumber",WareHouseOrderNumber,"Sales order summary for release to warehouse is Available")
Else
	Call AddResults(Iter,"D365_WareHouseOrderNumber",WareHouseOrderNumber,"Sales order summary for release to warehouse is Not Available")
End If
wait 2
Browser("All sales orders -- Finance").Sync
Call TakeScreenShot()

'Release the Order to Dallas
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebCheckBox("Select the current row").Click
wait 3
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebButton("Release to dallas").Click
wait 5
Call TakeScreenShot()

'Validate in Sales Order Details Again
'Click On Modules and wait for element to get loaded
Browser("Dashboard -- Finance and").Page("Dashboard -- Finance and").WebElement("Modules").Click
wait 2
'click on Accounts Receivable and wait for Page to Load
Browser("Dashboard -- Finance and").Page("Dashboard -- Finance and").WebElement("Accounts receivable").Click
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebButton("Collapse all").Click
wait 2
'click on Orders from Accounts Receivable Menu and wait for Page to Load
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Orders").Click
wait 2
'CLick on All Sales Order if Order is collapsed or Expanded
If Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("All sales orders").Exist(2) Then
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("All sales orders").Click
Else
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Orders").Click
	Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("All sales orders").Click
End If
wait 3
'Click on the Web Portal Number and enter the value as Order Number from Legacy Portal and Click on Apply Button
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("Web Portal Number").Click
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Sync
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebEdit("Filter field: Web Portal").Set DataTable.Value("LegacyOrderNumber","ResultSheet")
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Sync
wait 2
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("FIlter_Apply_Button").Click
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Sync
'Click on the Sales Order Number
Browser("All sales orders -- Finance").Page("All sales orders -- Finance").Link("SalesOrderNumber").Click
Browser("All sales orders -- Finance").Sync
'Click on Sales Order Header and Validate Values
'Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("SalesOrderHeader").Click
Browser("All sales orders -- Finance").Sync
wait 10
'Validate To Be Released Value
If Trim(Ucase(Browser("All sales orders -- Finance").Page("All sales orders -- Finance").WebElement("ToBeReleasedValue").GetROProperty("innertext"))) = "YES" Then
	Call AddResults(Iter,"D365_To Be Released Value","Yes","Pass")
Else
	Call AddResults(Iter,"D365_To Be Released Value","No","Fail")
	Environment.Value("FailureReason") = Environment.Value("FailureReason")  & "To Be Released Value is NO"
	Environment.Value("FailedApplication") = Environment.Value("FailedApplication") & "D365 Portal"
End  If
Call TakeScreenShot()

'Close the Browser
SystemUtil.CloseProcessByName("chrome.exe")










Function TakeScreenShot()
	On Error Resume Next
	wait 1
	ScreenShotFileName = Environment.Value("TestName") & "_" & year(now) & "_" & month(now) & "_" & day(now) & "_" & hour(now)&"_"&minute(now) & "_" & second(now)
	ScreenShotFilePath = Environment.Value("snapShotFolder") &"\"& ScreenShotFileName & ".png"
	Homepage.CaptureBitmap ScreenShotFilePath , False
End Function


Function AddResults(i,ColumnName,ColumnValue,Result)
	On Error Resume Next
	TotalRecord =  i+1
    DataTable.SetCurrentRow (TotalRecord+1)
    'create source column and update Values
    DataTable.Value(ColumnName,"ResultSheet") = ColumnValue
    If err.number <> 0 Then
        DataTable.GetSheet("ResultSheet").AddParameter ColumnName,""
    End If
        If err.number <> 0 Then
        DataTable.SetCurrentRow(TotalRecord+1)
        DataTable.Value(ColumnName,"ResultSheet") = ColumnValue
    End If
    'Update Field wise Result Value
    ResultColName = ColumnName & "-Result"
    DataTable.Value(ResultColName,"ResultSheet") = Ucase(Result)
    If err.number <> 0 Then
        DataTable.GetSheet("ResultSheet").AddParameter ResultColName,""
    End If
   
    If err.number <> 0 Then
    DataTable.SetCurrentRow(TotalRecord+1)
        DataTable.Value(ResultColName,"ResultSheet") = Ucase(Result)
    End If
   
   
End Function
