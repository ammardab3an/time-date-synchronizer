Option Explicit

Dim arrDateTime
Dim blnTest
Dim dtmDateTime, dtmNewDateTime
Dim intDateDiff, intOffset, intStatus, intThreshold, intTimeDiff
dim colItems, objHTTP, objItem, objRE, objWMIService
Dim strDateTime, strLocalDateTime, strMsg, strNewdateTime, strURL
dim ammarnewtime, ammarnewdate
dim oshell
' Defaults
intThreshold = 10
strURL       = "http://www.xs4all.nl/"

Set oShell = WScript.CreateObject ("WScript.Shell")
oShell.run ("cmd.exe /C date 01-01-2018")

' Check command line arguments
With WScript.Arguments
	If .Named.Count   > 0 Then Syntax
	If .Unnamed.Count > 2 Then Syntax
	If .Unnamed.Count > 0 Then
		If IsNumeric( .Unnamed(0) ) Then
			intThreshold = CInt( .Unnamed(0) )
		Else
			strURL = .Unnamed(0)
		End If
	End If
	If .Unnamed.Count = 2 Then
		If IsNumeric( .Unnamed(1) ) Then
			intThreshold = CInt( .Unnamed(1) )
		Else
			strURL = .Unnamed(1)
		End If
		' Only 1 argument should be numeric, not both
		If IsNumeric( .Unnamed(0) ) And IsNumeric( .Unnamed(1) ) Then
			Syntax
		End If
		' 1 argument should be numeric
		If Not ( IsNumeric( .Unnamed(0) ) Or IsNumeric( .Unnamed(1) ) ) Then
			Syntax
		End If
	End If
	' Threshold value must be between 0 and 60
	If intThreshold <  0 Then Syntax
	If intThreshold > 60 Then Syntax
	' URL must be a WEB server (full URL including protocol)
	blnTest = False
	Set objRE = New RegExp
	objRE.Pattern = "^https?://.+$"
	blnTest = objRE.Test( strURL )
	Set objRE = Nothing
	If Not blnTest Then Syntax
End With

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

' Abort if the server could not be reached
If intStatus <> 200 Then Syntax

' Convert the returned Apache timestamp string to a date to work with
arrDateTime = Split( strDateTime, " " )
strDateTime = arrDateTime(1) & " " _
            & arrDateTime(2) & " " _
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
	' Get current local system time ("before" time)
	strLocalDateTime = objItem.LocalDateTime
	' Add offset to GMT to get correct local time
	dtmNewDateTime   = DateAdd( "n", intOffset, dtmDateTime )
	' Format date and time string to be used to set new system time
	strNewdateTime   = Year( dtmNewDateTime ) _
	                 & Right( "0" & Month(  dtmNewDateTime ), 2 ) _
	                 & Right( "0" & Day(    dtmNewDateTime ), 2 ) _
	                 & Right( "0" & Hour(   dtmNewDateTime ), 2 ) _
	                 & Right( "0" & Minute( dtmNewDateTime ), 2 ) _
	                 & Right( "0" & Second( dtmNewDateTime ), 2 )
	
    ammarnewtime = Right( "0" & Hour(   dtmNewDateTime ), 2 ) _
					 & ":"_
	                 & Right( "0" & Minute( dtmNewDateTime ), 2 ) _
	                 & ":"_
					 & Right( "0" & Second( dtmNewDateTime ), 2 )
	
	oShell.run ("cmd.exe /C time " & ammarnewtime)
	
	ammarnewdate = 	 Right( "0" & Month(  dtmNewDateTime ), 2 ) _
					 & "-"_
	                 & Right( "0" & Day(    dtmNewDateTime ), 2 ) _
					 & "-"_
	                 & Year( dtmNewDateTime )
	
	oShell.run ("cmd.exe /C date " & ammarnewdate)
    

Next
Set colItems      = Nothing
Set objWMIService = Nothing


Sub Syntax( )
	msgbox "something go wrong"
    WScript.Quit 1
End Sub
msgbox "time and date have been updated (ammar dabaan)"