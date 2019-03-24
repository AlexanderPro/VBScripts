Option Explicit

call Main()

Function Main()
    Dim machineName, userName, password
    machineName = InputBox("Input machine IP or machine name:")
    userName = InputBox("Input user name:")
    password = InputBox("Input password:")
    WScript.Echo GetRemoteDateTime(machineName, userName, password)
End Function

Function GetRemoteDateTime(machineName, userName, password)
    Dim locator, wmi, item, itemCollection
    Const ns = "Root\CIMV2"
    Set locator = CreateObject("WbemScripting.SWbemLocator") 
    Set wmi = locator.ConnectServer(machineName, ns, userName, password)
    'Set wmi = locator.ConnectServer(machineName, ns)
    Set itemCollection = wmi.ExecQuery("SELECT * FROM Win32_LocalTime")
    For Each item In itemCollection
        GetRemoteDateTime = FormatDateTime(item.Day, item.Month, item.Year, item.Hour, item.Minute, item.Second)
    Next
End Function

Function PadLeft(s, ch, number)
    PadLeft = Right((String(number, ch) & s), number)
End Function

Function FormatDateTime(d, m, y, h, mi, s)
    FormatDateTime = PadLeft(d, "0", 2) & "." & PadLeft(m, "0", 2) & "." & PadLeft(y, "0", 4) & " " & PadLeft(h, "0", 2) & ":" & PadLeft(mi, "0", 2) & ":" & PadLeft(s, "0", 2)
End Function