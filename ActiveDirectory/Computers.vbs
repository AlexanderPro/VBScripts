Option Explicit

Const logFileName = "Computers.txt"

call Main()

Function Main()
    Dim defaultNamingContext, rootObject, element, elementCollection, text
    Set rootObject = GetObject("LDAP://rootDSE")
    'Set rootObject = rootObject.OpenDSObject("LDAP://" & domain & "/rootDSE", user, password, 1)
    defaultNamingContext = rootObject.Get("defaultNamingContext")
    Set elementCollection = GetObject("LDAP://CN=Computers, " & defaultNamingContext)
    text = "Domain contains next computers:" & vbCrLf & vbCrLf
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