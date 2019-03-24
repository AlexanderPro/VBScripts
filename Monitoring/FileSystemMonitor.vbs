Option Explicit

Const pathToLogFile = "FileSystemMonitor_{yyyy}{MM}{dd}.txt"
Const drive = "C:"
Const folder = "\\Temp\\"

call Main()

Function Main()
    Dim wmi, query, eventCollection, ev, targetInstance, PreviousInstance, prop, message
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
    query = "Select * From __InstanceOperationEvent Within 5 Where TargetInstance Isa 'CIM_DataFile' And TargetInstance.Drive='" & drive & "' And TargetInstance.Path='" & folder & "'"
    Set eventCollection = wmi.ExecNotificationQuery(query)
    Do
        Set ev = eventCollection.NextEvent()
        Set targetInstance = ev.TargetInstance
        Select Case ev.Path_.Class
            Case "__InstanceCreationEvent"
                message = Now & vbCrLf & vbCrLf & "Created:  " & targetInstance.Name & vbCrLf & "************************************************" & vbCrLf
            Case "__InstanceDeletionEvent"
                message = Now & vbCrLf & vbCrLf & "Deleted:  " & targetInstance.Name & vbCrLf & "************************************************" & vbCrLf
            Case "__InstanceModificationEvent"
                Set PreviousInstance = ev.PreviousInstance
                For Each prop in targetInstance.Properties_
                    If prop.Value <> PreviousInstance.Properties_(prop.Name) Then
                        message = Now & vbCrLf & vbCrLf & _
                                  "Changed:        " & targetInstance.Name                     & vbCrLf & _
                                  "Property:       " & prop.Name                               & vbCrLf & _
                                  "Previous value: " & PreviousInstance.Properties_(prop.Name) & vbCrLf & _
                                  "New value:      " & prop.Value                              & vbCrLf & _
                                  "************************************************"           & vbCrLf
                    End If
                Next
        End Select
        call WriteLineToLog(message)
    Loop
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

Function WriteLineToLog(text)
    Dim pathLog
    pathLog = FormatPath(pathToLogFile, Now())
    call WriteLineToFile(pathLog, text)
End Function