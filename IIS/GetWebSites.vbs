Option Explicit

Const logFileName = "WebSites.txt"

call Main()

Function Main()
    Dim service, server, text
    Const serverType     = "Web"       '"FTP"
    Const serverMetaType = "W3SVC"     '"MSFTPSVC"
    Set service = GetObject("IIS://localhost/" & serverMetaType)
    text = "Web server contains the sites:" & vbCrLf
    call WriteLineToFile(logFileName, text)
    For Each server in service
        If server.Class = "IIs" & serverType & "Server" Then
            text = "Id:                 " & server.Name & vbCrLf &_
                   "Name:               " & server.ServerComment & vbCrLf &_
                   "AppPoolId:          " & server.AppPoolId & vbCrLf &_
                   "State:              " & GetStateDescription(server.ServerState) & vbCrLf &_
                   "Connection timeout: " & server.ConnectionTimeout & vbCrLf &_
                   "HTTP Bindings:      " & EnumBindings(server.ServerBindings) & vbCrLf
            If serverType = "Web" Then
                text = text & "HTTPS Bindings:     " & EnumBindings(server.SecureBindings) & vbCrLf
            End If
            call WriteLineToFile(logFileName, text)
        End If
    Next
    WScript.Echo "Script is done!"
End Function

Function EnumBindings(bindingList)
    Dim i, ip, port, host
    Dim binding, match, matches
    Set binding = New RegExp
    binding.Pattern = "([^:]*):([^:]*):(.*)"
    For i = Lbound(bindingList) To Ubound(bindingList)
        ' bindingList( i ) is a string looking like IP:Port:Host
        Set matches = binding.Execute(bindingList(i))
        For Each match in matches
            ip = match.SubMatches(0)
            port = match.SubMatches(1)
            host = match.SubMatches(2)
            ' Do some pretty processing
            If ip = "" Then ip = "All Unassigned"
            If host = "" Then host = "*"
            If Len(ip) < 8 Then ip = ip & VbTab
            EnumBindings = EnumBindings & ip & VbTab & port & VbTab & host & VbTab & ""
        Next
    Next
End Function

Function GetStateDescription(state)
    Select Case state
    Case 1
        GetStateDescription = "Starting (MD_SERVER_STATE_STARTING)"
    Case 2
        GetStateDescription = "Started (MD_SERVER_STATE_STARTED)"
    Case 3
        GetStateDescription = "Stopping (MD_SERVER_STATE_STOPPING)"
    Case 4
        GetStateDescription = "Stopped (MD_SERVER_STATE_STOPPED)"
    Case 5
        GetStateDescription = "Pausing (MD_SERVER_STATE_PAUSING)"
    Case 6
        GetStateDescription = "Paused (MD_SERVER_STATE_PAUSED)"
    Case 7
        GetStateDescription = "Continuing (MD_SERVER_STATE_CONTINUING)"
    Case Else
        GetStateDescription = "Unknown state"
    End Select
End Function

Function WriteLineToFile(file, text)
    Dim fso, fout
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    Set fout = fso.OpenTextFile(file, 8, true)
    fout.WriteLine text
    fout.Close
End Function
