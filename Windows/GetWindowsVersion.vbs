Option Explicit

call Main()

Function Main
    Dim version
    version = GetWindowsVersion()
    version = StringOemToAnsi(version)
    WScript.Echo version
End Function

Function GetWindowsVersion
    Dim shell, exec, stream, res
    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.Exec("cmd /C ver")
    Set stream  = exec.StdOut
    res = ""
    Do While Not stream.AtEndOfStream
        res = res & stream.Read(1)
    Loop
    GetWindowsVersion = res
End Function

Function CharOemToAnsi(ch)
    Dim code, res
    code = Asc(ch)
    If ((code >= 128) and (code <= 175)) Then
        res = Chr(code + 64)
    ElseIf ((code >= 224) and (code <= 239)) Then
        res = Chr(code + 16)
    ElseIf (code = 241) Then
        res = Chr(184)
    ElseIf (code = 240) Then
        res = Chr(168)
    Else
        res = ch
    End If
    CharOemToAnsi = res
End Function

Function StringOemToAnsi(str)
    Dim i, res, length
    res = ""
    length = Len(str)
    For i = 1 To length Step 1
        res = res & CharOemToAnsi(Mid(str, i, 1))
    Next
    StringOemToAnsi = res
End Function