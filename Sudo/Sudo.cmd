@echo Set shell = CreateObject("Shell.Application") > %temp%\sudo.tmp.vbs
@echo args = Right("%*", (Len("%*") - Len("%1"))) >> %temp%\sudo.tmp.vbs
@echo shell.ShellExecute "%1", args, "", "runas", 1 >> %temp%\sudo.tmp.vbs
@start "" /B %temp%\sudo.tmp.vbs %*