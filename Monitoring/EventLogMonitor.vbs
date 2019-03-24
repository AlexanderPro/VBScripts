Option Explicit

Const pathToLogFile = "EventLogMonitor_{yyyy}{MM}{dd}.txt"

call Main()

Function Main()
    Dim wmi, eventCollection, ev, message
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
    Set eventCollection = wmi.ExecNotificationQuery("Select * from __InstanceCreationEvent within 5 where TargetInstance isa 'Win32_NTLogEvent'")
    Do
        Set ev = eventCollection.NextEvent
        message = Now & vbCrLf & vbCrLf & _
                  "Event ID:   " & ev.TargetInstance.EventCode                      & vbCrLf & _
                  "Event Type: " & ev.TargetInstance.Type                           & vbCrLf & _
                  "DateTime:   " & ConvertToDateTime(ev.TargetInstance.TimeWritten) & vbCrLf & _
                  "Source:     " & ev.TargetInstance.SourceName                     & vbCrLf & _
                  "Category:   " & ev.TargetInstance.CategoryString                 & vbCrLf & _
                  "User:       " & ev.TargetInstance.User                           & vbCrLf & _
                  "Computer:   " & ev.TargetInstance.ComputerName                   & vbCrLf & _
                  "Log File:   " & ev.TargetInstance.Logfile                        & vbCrLf & _
                  "Text:       " & ev.TargetInstance.Message                        & vbCrLf & _
                  "******************************************************************"    & vbCrLf & vbCrLf
        call WriteLineToLog(message)
    Loop
End Function

'Formats date from WMI type to VBScript type
Function ConvertToDateTime(dateTime)
    ConvertToDateTime = CDate(Mid(dateTime, 7, 2) & "." & Mid(dateTime, 5, 2) & "." & Mid(dateTime, 1, 4) & " " & Mid(dateTime, 9, 2) & ":" & Mid(dateTime, 11, 2) & ":" & Mid(dateTime, 13, 2))
End Function

Function FormatPath(path, dateTime)
    Dim dd, mm, mmmm, yyyy
    dd = Day(dateTime)
    dd = "0" & dd
    dd = Right(dd, 2)
    mm = Month(dateTime)
    mm = "0" & mm
    mm = Right(mm, 2)
    yyyy = Year(dateTime)
    yyyy = "000" & yyyy
    yyyy = Right(yyyy, 4)
    mmmm = MonthName(Month(dateTime))
    FormatPath = path
    FormatPath = Replace(FormatPath, "{yyyy}", yyyy)
    FormatPath = Replace(FormatPath, "{MM}", mm)
    FormatPath = Replace(FormatPath, "{MMMM}", mmmm)
    FormatPath = Replace(FormatPath, "{dd}", dd)
End Function

Function WriteLineToFile(file, text)
    Dim fso, fout
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    Set fout = fso.OpenTextFile(file, 8, true)
    fout.WriteLine text
    fout.Close
End Function

Function WriteLineToLog(text)
    Dim pathLog
    pathLog = FormatPath(pathToLogFile, Now())
    call WriteLineToFile(pathLog, text)
End Function