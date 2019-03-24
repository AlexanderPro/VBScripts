Option Explicit

Const domainName = "WORKGROUP"
Const logFileName = "Users.txt"

call Main()

Function Main
    Dim ns, element, elementCollection, text
    ns = "WinNT://" & domainName
    Set elementCollection = GetObject(ns)
    elementCollection.Filter = Array("User")
    text = "Domain contains the users:" & vbCrLf & vbCrLf
    For Each element In elementCollection
        text = text & element.Name & ";" & element.FullName & vbCrLf
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