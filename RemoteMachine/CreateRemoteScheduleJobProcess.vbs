Option Explicit

call Main()

Function Main()
    Dim machineName, user, password, processName, err
    machineName = InputBox("Input machine IP or machine name:")
    user = InputBox("Input user name:")
    password = InputBox("Input password:")
    processName = InputBox("Input process name or process path:") 
    err = CreateProcessRemotely(machineName, user, password, processName)
    If err = 0 Then
        WScript.Echo "Scheduled Job is started." & vbCrLf & "Script is done!"
    Else
        WScript.Echo "Process could not be started due to error: " & err & vbCrLf & "Script is done!"
    End If
End Function

Function CreateProcessRemotely(machineName, user, password, processName)
    Dim locator, wmi, dateTime, dateTimeCollection, scheduledJob, swbemDateTime, jobID, remoteDateTime
    Const ns = "Root\CIMV2"
    Set locator = CreateObject("WbemScripting.SWbemLocator")
    'Set wmi = locator.ConnectServer(machineName, ns, user, password, "MS_409", "ntlmdomain:" + strDomain) 
    'Set wmi = locator.ConnectServer(machineName, ns)
    Set wmi = locator.ConnectServer(machineName, ns, user, password)
    Set dateTimeCollection = wmi.ExecQuery("Select * From Win32_LocalTime")
    For Each dateTime In dateTimeCollection
        remoteDateTime = CreateDateTime(dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour, dateTime.Minute, dateTime.Second)
    Next
    remoteDateTime = DateAdd("n", 1, remoteDateTime)
    Set scheduledJob = wmi.Get("Win32_ScheduledJob")
    Set swbemDateTime = CreateObject("WbemScripting.SWbemDateTime")
    swbemDateTime.SetVarDate(remoteDateTime)
    CreateProcessRemotely = scheduledJob.Create(processName, swbemDateTime.Value, False, 0, 0, True, jobID)
End Function

Function CreateDateTime(year, month, day, hour, minute, second)
    Dim dt
    dt = DateSerial(year, month, day)
    dt = DateAdd("h", hour, dt)
    dt = DateAdd("n", minute, dt)
    dt = DateAdd("s", second, dt)
    CreateDateTime = dt
End Function