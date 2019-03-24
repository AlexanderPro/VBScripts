Option Explicit

Const logFileName = "Domains.txt"

call Main()

Function Main()
    Dim element, elementCollection, text
    Set elementCollection = GetObject("WinNT:")
    text = "List of domains:" & vbCrLf & vbCrLf
    For Each element In elementCollection
        text = text & element.Name & vbCrLf
    Next
    call WriteLineToFile(logFileName, text)
    WScript.Echo "Script is done!"
End Function

Function WriteLineToFile(file, text)
    Dim fso, fout
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    Set fout = fso.OpenTextFile(file, 8, true)
    fout.WriteLine text
    fout.Close
End Function