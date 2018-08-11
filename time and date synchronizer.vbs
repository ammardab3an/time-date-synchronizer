Option Explicit

Dim arrDateTime
Dim dtmDateTime, dtmNewDateTime
Dim intOffset, intStatus
dim colItems, objHTTP, objItem, objWMIService
Dim strDateTime, strURL
dim dtMonth
dim ammarnewtime, ammarnewdate, aa, bb
dim  oShell
dim theFirstMsgbox, theSecondMsgbox, theInputBox

' Defaults
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
theFirstMsgbox = MsgBox ("if the script show 'CDate' error then chose Yes, if this is your first time you run it chose No", vbYesNo, "Ammar Dab3an")

Select Case theFirstMsgbox
Case vbYes
    
dtMonth = monthConvert(arrDateTime(2))
strDateTime = arrDateTime(1) & " " _
            & dtMonth & " " _
            & arrDateTime(3) & " " _
            & arrDateTime(4)

Case vbNo

strDateTime = arrDateTime(1) & " " _
            & arrDateTime(2) & " " _
            & arrDateTime(3) & " " _
            & arrDateTime(4)

End Select


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
	
theSecondMsgbox = MsgBox ("Do you want to git your timezone form your system", vbYesNo, "Ammar Dab3an")

Select Case theSecondMsgbox
Case vbYes
	' Get timezone offset telative to GMT
	intOffset        = CInt( objItem.CurrentTimeZone )
Case vbNo
theInputBox  = InputBox ( "Enter how many minutes do you want to add to the current GMT time, please type only numbers like '60' for an hour or '180' for three hours", "Ammar Dab3an" )
intOffset        = CInt(theInputBox)
End Select


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
	
	
 
	
	aa =  ("/C date " & ammarnewdate)
	bb = ("/C time " & ammarnewtime)
	

	oShell.ShellExecute "cmd.exe",aa , , "runas", 1
	oShell.ShellExecute "cmd.exe",bb , , "runas", 1
	
Next
Set colItems      = Nothing
Set objWMIService = Nothing


Sub Syntax( )
	msgbox "some thing go wrong, check the internet connection"
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

msgbox "your time and date have been updated (made by Ammar Dab3an)"