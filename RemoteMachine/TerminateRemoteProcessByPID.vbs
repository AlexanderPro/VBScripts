Option Explicit

call Main()

Function Main
    Dim machineName, user, password, pid, status
    machineName = InputBox("Input machine IP or machine name:")
    user = InputBox("Input user name:")
    password = InputBox("Input password:")
    pid = InputBox("Input PID:")
    status = TerminateProcessRemotely(machineName, user, password, pid)
    WScript.Echo "Process was terminated with status: " & status & VbCrLf & "Script is done!"
End Function

Function TerminateProcessRemotely(machineName, user, password, pid)
    Dim locator, wmi, process, processCollection
    Const ns = "Root\CIMV2"
    Set locator = CreateObject("WbemScripting.SWbemLocator") 
    'Set wmi = locator.ConnectServer(machineName, ns, user, password, "MS_409", "ntlmdomain:" + strDomain)
    'Set wmi = locator.ConnectServer(machineName, ns)
    Set wmi = locator.ConnectServer(machineName, ns, user, password)
    Set processCollection = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = '" & pid & "'")
    TerminateProcessRemotely = -1
    For Each process In processCollection
        TerminateProcessRemotely = process.Terminate()
    Next
End Function