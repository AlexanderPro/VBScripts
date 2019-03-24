Option Explicit

Const sourceFilePath = "Windows.txt"
Const destinationFilePath = "Dos.txt"

call Main()

Function Main()
    call FileAnsiToOem(sourceFilePath, destinationFilePath)
    WScript.Echo "Script is done!"
End Function

Function FileAnsiToOem(sourceFile, destinationFile)
    Dim fso, fRead, fWrite, ch
    Const ForReading = 1, ForWriting = 2
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fRead = fso.OpenTextFile(sourceFile, ForReading, False)
    Set fWrite = fso.OpenTextFile(destinationFile, ForWriting, True)
    Do While fRead.AtEndOfStream <> True
        ch = fRead.Read(1)
        ch = CharAnsiToOem(ch)
        fWrite.Write(ch)
    Loop
    fRead.Close()
    fWrite.Close()
End Function

Function CharAnsiToOem(ch)
    Dim code, res
    code = Asc(ch)
    If ((code >= 192) and (code <= 239)) Then
        res = Chr(code - 64)
    ElseIf ((code >= 240) and (code <= 255)) Then
        res = Chr(code - 16)
    ElseIf (code = 184) Then
        res = Chr(241)
    ElseIf (code = 168) Then
        res = Chr(240)
    Else
        res = ch
    End If
    CharAnsiToOem = res
End Function

Function StringAnsiToOem(str)
    Dim i, res, length
    res = ""
    length = Len(str)
    For i = 1 To length Step 1
        res = res & CharAnsiToOem(Mid(str, i, 1))
    Next
    StringAnsiToOem = res
End Function