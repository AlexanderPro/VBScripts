Option Explicit

Const logFileName = "Departments.txt"

call Main()

Function Main()
    Dim defaultNamingContext, dictionary, rootObject, element, elementCollection
    Set rootObject = GetObject("LDAP://rootDSE")
    'Set rootObject = rootObject.OpenDSObject("LDAP://" & domain & "/rootDSE", user, password, 1)
    defaultNamingContext = rootObject.Get("defaultNamingContext")
    Set elementCollection = GetObject("LDAP://CN=Users, " & defaultNamingContext)
    Set dictionary = CreateObject("Scripting.Dictionary")
    For Each element In elementCollection
        If Not dictionary.Exists(element.department) Then
            dictionary.Add element.department, ""
        End If
    Next
    call WriteLineToFile(logFileName, "Domain contains next departments:")
    call WriteDictionaryToLog(dictionary)
    WScript.Echo "Script is done!"
End Function

Function WriteDictionaryToLog(dictionary)
    Dim keys, i
    keys = SortArrayAsc(dictionary.Keys)
    For i = 0 To dictionary.Count - 1
        call WriteLineToFile(logFileName, keys(i))
    Next
End Function

Function WriteLineToFile(file, text)
    Dim fso, fout
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    Set fout = fso.OpenTextFile(file, 8, true)
    fout.WriteLine text
    fout.Close
End Function

Function SortArrayAsc(a)
    Dim i, j, temp
    For i = UBound(a) - 1 To 0 Step -1
        For j= 0 to i
            If a(j) > a(j+1) Then
                temp = a(j+1)
                a(j+1) = a(j)
                a(j) = temp
            End If
        Next
    Next 
    SortArrayAsc = a
End Function