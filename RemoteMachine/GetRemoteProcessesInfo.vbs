Option Explicit

Const pathToLogFile = "Processes_{yyyy}{MM}{dd}.txt"

call Main()

Function Main()
    Dim machineName, user, password, locator, wmi, process, processCollection, message, processNumber
    Const ns = "Root\CIMV2"
    machineName = InputBox("Input machine IP or machine name:")
    user = InputBox("Input user name:")
    password = InputBox("Input password:")
    Set locator = CreateObject("WbemScripting.SWbemLocator") 
    'Set wmi = locator.ConnectServer(machineName, ns, user, password, "MS_409", "ntlmdomain:" + strDomain)
    'Set wmi = locator.ConnectServer(machineName, ns)
    Set wmi = locator.ConnectServer(machineName, ns, user, password)
    Set processCollection = wmi.ExecQuery("SELECT * FROM Win32_Process")
    message = "  ¹" & String(5, " ") & "Name" & String(33, " ") & "PID" & String(19, " ") & "Start DateTime" & String(14, " ") & "Threads" & String(14, " ") & "WorkingSetSize" & String(14, " ") & "Executable Path"
    call WriteLineToLog(message)
    
    processNumber = 0
    For Each process In processCollection
        processNumber = processNumber + 1
        message = PadLeft(processNumber, 3) & String(5, " ") &_ 
                  PadRight(process.Name, 30) &_ 
                  PadLeft(process.ProcessId, 10) &_
                  PadLeft(ConvertToDateTime(process.CreationDate), 33) &_
                  PadLeft(process.ThreadCount, 21) &_
                  PadLeft(FormatBytes(process.WorkingSetSize), 28) & String(14, " ") &_
                  process.ExecutablePath
        call WriteLineToLog(message)
    Next
    WScript.Echo "Script is done!"
End Function

Function PadLeft(s, number)
    PadLeft = Right((String(number, " ") & s), number)
End Function

Function PadRight(s, number)
    PadRight = Left((s & String(number, " ")), number)
End Function

Function FormatBytes(number)
    FormatBytes = FormatNumber((number / 1024), 2) & " KB"
End Function

Function ConvertToDateTime(dateTime)
    If IsNull(dateTime) Then
        ConvertToDateTime = "N/A"
    Else
        ConvertToDateTime = CDate(Mid(dateTime, 7, 2) & "." & Mid(dateTime, 5, 2) & "." & Mid(dateTime, 1, 4) & " " & Mid(dateTime, 9, 2) & ":" & Mid(dateTime, 11, 2) & ":" & Mid(dateTime, 13, 2))
    End If
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