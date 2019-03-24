Option Explicit
Dim args, shell, str, i

Set args = WScript.Arguments
If args.Count < 1 Then
    WScript.Echo "Usage: sudo <arg1 arg2 .. argN>"
    WScript.Quit(-1)
End If
For i = 1 to args.Count - 1
    str = str + " " + args(i)
Next
Set shell = CreateObject("Shell.Application") 
shell.ShellExecute args(0), str, "", "runas", 1
WScript.Quit(0)