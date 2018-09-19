Option Explicit

dim oShell
Dim arrDateTime, dtmDateTime, tmpDateTime
dim colItems, objHTTP, objItem, objWMIService
Dim strDateTime, strURL
dim dtMonth, intOffset
dim newDate, newTime

' Defaults
strURL       = "http://time.windows.com/"

Set oShell = WScript.CreateObject ("Shell.Application")
oShell.ShellExecute  "cmd.exe", "/c date 09-09-2018" , , "runas", 1

' Get server time from a web server
Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
objHTTP.Open "GET", strURL, False
objHTTP.SetRequestHeader "User-Agent", WScript.ScriptName
On Error Resume Next
objHTTP.Send

strDateTime = objHTTP.GetResponseHeader( "Date" )
Set objHTTP = Nothing
If Err Then Syntax

' Convert the returned Apache timestamp string to a date to work with
arrDateTime = Split( strDateTime, " " )
    
dtMonth = monthConvert(arrDateTime(2))
strDateTime = arrDateTime(1) & " " _
            & dtMonth & " " _
            & arrDateTime(3) & " " _
            & arrDateTime(4)

dtmDateTime = CDate( strDateTime )

' Get and set local system date and time
Set objWMIService = GetObject( "winmgmts:{(Systemtime)}//./root/CIMV2" )
Set colItems      = objWMIService.ExecQuery( "Select * From Win32_OperatingSystem" )
For Each objItem In colItems
	
	' Get timezone offset telative to GMT
	intOffset = CInt( objItem.CurrentTimeZone )

	' Add offset to GMT to get correct local time
	tmpDateTime = DateAdd( "n", intOffset, dtmDateTime )
	
    newTime = Right( "0" & Hour(   tmpDateTime ), 2 ) & ":" & _
			  Right( "0" & Minute( tmpDateTime ), 2 ) & ":" & _
			  Right( "0" & Second( tmpDateTime ), 2 )
	
		
	newDate = Right( "0" & Month(  tmpDateTime ), 2 ) & "-" & _
			  Right( "0" & Day(    tmpDateTime ), 2 )  & "-" & _
			  Year( tmpDateTime )

	oShell.ShellExecute "cmd.exe", ("/C date " & newDate), , "runas", 1
	oShell.ShellExecute "cmd.exe", ("/C time " & newTime), , "runas", 1
	
Next

msgbox "your time and date have been updated (Ammar Dab3an)"

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


