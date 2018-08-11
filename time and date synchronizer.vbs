Option Explicit

Dim arrDateTime
Dim blnTest
Dim dtmDateTime, dtmNewDateTime
Dim intDateDiff, intOffset, intStatus, intThreshold, intTimeDiff
dim colItems, objHTTP, objItem, objRE, objWMIService
Dim strDateTime, strLocalDateTime, strMsg, strNewdateTime, strURL
dim dtMonth
dim ammarnewtime, ammarnewdate, aa, bb
dim  oShell, ashell, bshell
' Defaults
intThreshold = 10
strURL       = "http://time.windows.com/"

Set oShell = WScript.CreateObject ("Shell.Application")
oShell.ShellExecute  "cmd.exe", "/c date 08-08-2018" , , "runas", 1

' Get server time from a web server
Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
objHTTP.Open "GET", strURL, False
objHTTP.SetRequestHeader "User-Agent", WScript.ScriptName
On Error Resume Next
objHTTP.Send
intStatus   = objHTTP.Status
strDateTime = objHTTP.GetResponseHeader( "Date" )
Set objHTTP = Nothing
If Err Then Syntax
On Error Goto 0



' Convert the returned Apache timestamp string to a date to work with
arrDateTime = Split( strDateTime, " " )
dtMonth = monthConvert(arrDateTime(2))
strDateTime = arrDateTime(1) & " " _
            & dtMonth & " " _
            & arrDateTime(3) & " " _
            & arrDateTime(4)
dtmDateTime = CDate( strDateTime )
strDateTime = Year( dtmDateTime ) _
            & Right( "0" & Month(  dtmDateTime ), 2 ) _
            & Right( "0" & Day(    dtmDateTime ), 2 ) _
            & Right( "0" & Hour(   dtmDateTime ), 2 ) _
            & Right( "0" & Minute( dtmDateTime ), 2 ) _
            & Right( "0" & Second( dtmDateTime ), 2 )

' Get and set local system date and time
Set objWMIService = GetObject( "winmgmts:{(Systemtime)}//./root/CIMV2" )
Set colItems      = objWMIService.ExecQuery( "Select * From Win32_OperatingSystem" )
For Each objItem In colItems
	' Get timezone offset telative to GMT
	intOffset        = CInt( objItem.CurrentTimeZone )
	msgbox  (intoffset)
	' Add offset to GMT to get correct local time
	dtmNewDateTime   = DateAdd( "n", intOffset, dtmDateTime )
	
    ammarnewtime = Right( "0" & Hour(   dtmNewDateTime ), 2 ) _
					 & ":"_
	                 & Right( "0" & Minute( dtmNewDateTime ), 2 ) _
	                 & ":"_
					 & Right( "0" & Second( dtmNewDateTime ), 2 )
	
	
	
	ammarnewdate = 	 Right( "0" & Month(  dtmNewDateTime ), 2 ) _
					 & "-"_
	                 & Right( "0" & Day(    dtmNewDateTime ), 2 ) _
					 & "-"_
	                 & Year( dtmNewDateTime )
	
	
    
	msgbox (ammarnewtime)
	msgbox (ammarnewdate)
	
	aa =  ("/C date " & ammarnewdate)
	bb = ("/C time " & ammarnewtime)
	
	Set aShell = CreateObject("Shell.Application")
	Set bShell = CreateObject("Shell.Application")
	aShell.ShellExecute "cmd.exe",aa , , "runas", 1
	bShell.ShellExecute "cmd.exe",bb , , "runas", 1
	
Next
Set colItems      = Nothing
Set objWMIService = Nothing


Sub Syntax( )
	msgbox "some thing go wrong, check the internet connection or make the cmd run as administrator"
    WScript.Quit 1
End Sub

function  monthConvert(month)

select case month 
case "Jan"
monthConvert = "01"
case "Feb"
monthConvert = "02"
case "Mar"
monthConvert = "03"
case "Apr"
monthConvert = "04"
case "May"
monthConvert = "05"
case "Jun"
monthConvert = "06"
case "Jul"
monthConvert = "07"
case "Aug"
monthConvert = "08"
case "Sep"
monthConvert = "09"
case "Oct"
monthConvert = "10"
case "Nov"
monthConvert = "11"
case "Dec"
monthConvert = "12"

case "January"
monthConvert = "01"
case "February"
monthConvert = "02"
case "March"
monthConvert = "03"
case "April"
monthConvert = "04"
case "June"
monthConvert = "06"
case "July"
monthConvert = "07"
case "August"
monthConvert = "08"
case "September"
monthConvert = "09"
case "October"
monthConvert = "10"
case "November"
monthConvert = "11"
case "December"
monthConvert = "12"

end select

end function

msgbox "time and date have been updated (ammar dabaan)"