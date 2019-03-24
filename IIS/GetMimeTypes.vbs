Option Explicit

Const pathToLogFile = "MimeTypes.txt"

call Main()

Function Main
    Dim mimeMap, mimeMapObject, message
    Set mimeMapObject = GetObject("IIS://LocalHost/MimeMap")
    For Each mimeMap in mimeMapObject.GetEx("MimeMap")
        message = message & mimeMap.Extension & "  " & mimeMap.MimeType & VbCrLf
    Next
    call WriteLineToFile(pathToLogFile, message)
    WScript.Echo("Script is done!")
End Function

Function FormatPath(path, dateTime)
    Dim dd, mm, mmmm, yyyy
    dd = Day(dateTime)
    dd = "0" & dd
    dd = Right(dd, 2)
    mm = Month(dateTime)
    mm = "0" & mm
    mm = Right(mm, 2)
    yyyy = Year(dateTime)
    yyyy = "000" & yyyy
    yyyy = Right(yyyy, 4)
    mmmm = MonthName(Month(dateTime))
    FormatPath = path
    FormatPath = Replace(FormatPath, "{yyyy}", yyyy)
    FormatPath = Replace(FormatPath, "{MM}", mm)
    FormatPath = Replace(FormatPath, "{MMMM}", mmmm)
    FormatPath = Replace(FormatPath, "{dd}", dd)
End Function

Function WriteLineToFile(file, text)
    Dim fso, fout
    Set fso = WScript.CreateObject("Scripting.FileSystemObject")
    Set fout = fso.OpenTextFile(file, 8, true)
    fout.WriteLine text
    fout.Close
End Function