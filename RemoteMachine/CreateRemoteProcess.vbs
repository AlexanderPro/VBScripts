Option Explicit

call Main()

Function Main()
    Dim machineName, user, password, processName, processID
    machineName = InputBox("Input machine IP or machine name:")
    user = InputBox("Input user name:")
    password = InputBox("Input password:")
    processName = InputBox("Input process name or process path:")
    processID = CreateProcessRemotely(machineName, user, password, processName)
    WScript.Echo "Process was started with PID: " & processID & vbCrLf & "Script is done!"
End Function

Function CreateProcessRemotely(machineName, user, password, processName)
    Dim locator, wmi, item, itemCollection, startup, process, config, processID
    Const ns = "Root\CIMV2"
    Set locator = CreateObject("WbemScripting.SWbemLocator")
    'Set objWMI = locator.ConnectServer(machineName, ns, strUser, strPassword, "MS_409", "ntlmdomain:" + strDomain) 
    'Set wmi = locator.ConnectServer(machineName, ns)
    Set wmi = locator.ConnectServer(machineName, ns, user, password)
    Set startup = wmi.Get("Win32_ProcessStartup")
    Set config = startup.SpawnInstance_
    config.ShowWindow = 10
    Set process = wmi.Get("Win32_Process")
    process.Create processName, Null, config, processID
    CreateProcessRemotely = processID
End Function