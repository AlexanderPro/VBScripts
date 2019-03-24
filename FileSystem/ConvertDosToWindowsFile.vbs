Option Explicit

Const sourceFilePath = "Dos.txt"
Const destinationFilePath = "Windows.txt"

call Main()

Function Main()
    call FileOemToAnsi(sourceFilePath, destinationFilePath)
    WScript.Echo "Script is done!"
End Function

Function FileOemToAnsi(sourceFile, destinationFile)
    Dim fso, fRead, fWrite, ch
    Const ForReading = 1, ForWriting = 2
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fRead = fso.OpenTextFile(sourceFile, ForReading, False)
    Set fWrite = fso.OpenTextFile(destinationFile, ForWriting, True)
    Do While fRead.AtEndOfStream <> True
        ch = fRead.Read(1)
        ch = CharOemToAnsi(ch)
        fWrite.Write(ch)
    Loop
    fRead.Close()
    fWrite.Close()
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