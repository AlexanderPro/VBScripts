Option Explicit

call Main()

Function Main()
    Dim machineName, user, password, result
    machineName = InputBox("Input machine IP or machine name:")
    user = InputBox("Input user name:")
    password = InputBox("Input password:")
    result = MsgBox("Reboot computer?", vbYesNo, "Visual Basic script")
    If result = vbYes Then
        call ExitWindows(machineName, user, password, True)
    else
        result = MsgBox("Shutdown computer?", vbYesNo, "Visual Basic script")
        If result = vbYes Then
            call ExitWindows(machineName, user, password, False)
        end If
    end If
    WScript.Echo "Script is done!"
End Function

Function ExitWindows(machineName, user, password, reboot)
    Dim locator, wmi, system, systemCollection
    Const ns = "Root\CIMV2"
    Set locator = CreateObject("WbemScripting.SWbemLocator") 
    'Set wmi = locator.ConnectServer(machineName, ns, user, password, "MS_409", "ntlmdomain:WG_FORS-BS_RZN")
    'Set wmi = locator.ConnectServer(machineName, ns)
    Set wmi = locator.ConnectServer(machineName, ns, user, password)
    wmi.Security_.ImpersonationLevel = 3
    wmi.Security_.Privileges.Add 23
    Set systemCollection = wmi.ExecQuery("Select Name From Win32_OperatingSystem")
    For Each system In systemCollection
        If reboot = True Then
            system.Reboot
        else
            system.Shutdown
        end If
    Next
End Function