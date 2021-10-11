'*************************'Demo Test Case for Metcash''*************************
'*************************'Designed By Koventhan '*************************
'the data will be used from UFT data Tables. Report also will be generated same.
'*************************'Pre Requisites :'*************************
'	1. All the Values will be taken from the input data sheet. Make Sure you give all the login details properly
'	2. Input data needs to be given in Data sheet. Make sure proper values are passed.
'Create Output Results Sheet

SystemUtil.Run "C:\Program Files\TUN\exe\daltest.exe"
wait 30
'msgbox(DataTable.Value("Dallas_UserName","Sheet1"))
Set Homepage =  Window("daltest").Window("daltest1")
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type "gmko"
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type "abcd1234"
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 2
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 250
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micF3 
wait 3
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type "ISORA"
wait 5
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 5
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micRight 
wait 5
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 3
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micRight 
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 3
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type "R"
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn 
wait 2
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micF3 @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf2.xml_;_
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type "ISORB" @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf3.xml_;_
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf4.xml_;_
wait 2
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf5.xml_;_
wait 2
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micRight @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf6.xml_;_
wait 2
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf7.xml_;_
wait 2
Call TakeScreenShot()
Window("daltest").Window("daltest1").WinObject("AfxFrameOrView40").Type  micReturn @@ hightlight id_;_593082_;_script infofile_;_ZIP::ssf8.xml_;_
wait 2
Call TakeScreenShot()



SystemUtil.CloseProcessByName("daltest.exe")



Function TakeScreenShot()
	On Error Resume Next
	wait 1
	ScreenShotFileName = Environment.Value("TestName") & "_" & year(now) & "_" & month(now) & "_" & day(now) & "_" & hour(now)&"_"&minute(now) & "_" & second(now)
	ScreenShotFilePath = Environment.Value("snapShotFolder") &"\"& ScreenShotFileName & ".png"
	Homepage.CaptureBitmap ScreenShotFilePath , False
End Function
