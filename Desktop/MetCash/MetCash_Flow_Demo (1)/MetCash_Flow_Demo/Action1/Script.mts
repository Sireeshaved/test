'*************************'Demo Test Case for Metcash''*************************
'*************************'Designed By Koventhan '*************************
'the data will be used from UFT data Tables. Report also will be generated same.
'*************************'Pre Requisites :'*************************
'	1. All the Values will be taken from the input data sheet. Make Sure you give all the login details properly
'	2. Input data needs to be given in Data sheet. Make sure proper values are passed.
'Create Output Results Sheet
On Error Resume Next
gbReportFilePath= "C:\Users\gmko\Desktop\TestExecution"
TestCaseFolder = gbReportFilePath & "\" & Environment.Value("TestName")
TestResultFolder = TestCaseFolder & "\" & year(now)&"_"&month(now)&"_"&day(now)&"_"&hour(now)&"_"&minute(now)
snapShotFolder = TestResultFolder & "\SnapShot"
Environment.Value("snapShotFolder") = snapShotFolder
Environment.Value("FailureReason") = ""
Environment.Value("FailedApplication") = ""

Iter = Environment.Value("TestIteration")-2
tcName = "Result_" & Environment.Value("TestName")& "_" & month(now) & "_" & year(now) & "_" & hour(now) & "_" & minute(now)&"_"&second(now)
Call CreateResultsFolders()
DataTable.AddSheet("ResultSheet")
Set Homepage =  Browser("Tasman Liquor Company")
'import the data sheet
DataTable.AddSheet("Sheet1")
DataTable.ImportSheet "C:\Users\gmko\Desktop\TestExecution\Demo_TestCase.xlsx","Sheet1","Sheet1"
'Launch the Chrome Browser

'For i =1 to  DataTable.GetSheet("Sheet1").GetRowCount
'DataTable.SetCurrentRow(i)
'Next
' Load the Legacy URL
SystemUtil.Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" ,DataTable.Value("Legacy_URL","Sheet1")
wait 30
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").WebEdit("CustomerId").Set DataTable.Value("CustomerID","Sheet1")
Browser("Tasman Liquor Company").Sync
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").WebEdit("Password").Set DataTable.Value("Password","Sheet1")
Browser("Tasman Liquor Company").Sync
Call TakeScreenShot()
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Image("b_login").Click 
wait 5
Browser("Tasman Liquor Company").Sync
'Choose Quick Order
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("nav").Image("QuickOrder").Click
Browser("Tasman Liquor Company").Sync
Call TakeScreenShot()
'add products to cart
If DataTable.Value("Product_Code1","Sheet1") <> "" Then
	Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").WebEdit("ProductId1").Set DataTable.Value("Product_Code1","Sheet1")
	Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").WebEdit("OrderCtnQty1").Set DataTable.Value("Qunatity1","Sheet1")
	wait 2
End If
If DataTable.Value("Product_Code2","Sheet1") <> "" Then
	Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").WebEdit("ProductId2").Set DataTable.Value("Product_Code2","Sheet1")
	Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").WebEdit("OrderCtnQty2").Set DataTable.Value("Qunatity2","Sheet1")
	wait 2	
End If
Browser("Tasman Liquor Company").Sync
Call TakeScreenShot()
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").Image("Add_To_Basket").Click
'click on View Order Basket
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("summary").Image("ViewOrderBasket").Click
Browser("Tasman Liquor Company").Sync
'Click on Place or Confirm Order
Call TakeScreenShot()
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").Image("PlaceOrConfirmOrder").Click
Browser("Tasman Liquor Company").Sync
'Add Customer OrderNumber and CLick on Send Order
Call TakeScreenShot()
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").WebEdit("CustomerRefNumber").Set DataTable.Value("CustomerRefNumber","Sheet1")
wait 2
Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").Image("SendOrder").Click
Call TakeScreenShot()
'Capture the OrderNumber
OrderDetail = Browser("Tasman Liquor Company").Page("Tasman Liquor Company").Frame("content").WebElement("Order Number").GetROProperty("innerhtml")
OrderDetails = Split(OrderDetail,":")
LegacyOrderNumber = Trim(OrderDetails(1))
'Add the Captured Value into result Sheet
If LegacyOrderNumber <>  "" Then
	Call AddResults(Iter,"LegacyOrderNumber",LegacyOrderNumber,"Order Number Generated")
Else
	Call AddResults(Iter,"LegacyOrderNumber",LegacyOrderNumber,"Order Number Not Generated")
	Environment.Value("FailureReason") = Environment.Value("FailureReason") & "Legacy Order Number Not Generated"
	Environment.Value("FailedApplication") =  Environment.Value("FailedApplication") & "Legacy Portal"
End If
'Close the Browser
SystemUtil.CloseProcessByName("chrome.exe")


'Call D365 Application Code
RunAction "D365Application", oneIteration

'Call Dallas Application Code
RunAction "DallasApplication", oneIteration


'Import The Results Sheet
Call AddResults(Iter,"FailureReasons",Environment.Value("FailureReason"),"")
Call AddResults(Iter,"FailedApplications",Environment.Value("FailedApplication"),"")
DataTable.ExportSheet TestResultFolder &"\"& tcName &".xlsx" , "ResultSheet" , "ResultSheet"




Function CreateResultsFolders()
	Dim fso, folder
	Set fso = CreateObject("Scripting.FileSystemObject")

	If Not (fso.FolderExists(TestCaseFolder)) Then
		fso.CreateFolder(TestCaseFolder)
	End If
	
	If Not (fso.FolderExists(TestResultFolder)) Then
		fso.CreateFolder(TestResultFolder)
	End If

	If Not (fso.FolderExists(snapShotFolder)) Then
		fso.CreateFolder(snapShotFolder)
	End If

End  Function

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


Function TakeScreenShot()
	On Error Resume Next
	wait 1
	ScreenShotFileName = Environment.Value("TestName") & "_" & year(now) & "_" & month(now) & "_" & day(now) & "_" & hour(now)&"_"&minute(now) & "_" & second(now)
	ScreenShotFilePath = Environment.Value("snapShotFolder") &"\"& ScreenShotFileName & ".png"
	Homepage.CaptureBitmap ScreenShotFilePath , False
End Function

