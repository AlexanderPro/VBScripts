Option Explicit

call Main()

Function Main
    Dim shell
    Set shell = WScript.CreateObject("WScript.Shell")

    '1. Input password on start screen with default English keyboard
    shell.RegWrite "HKEY_USERS\.DEFAULT\Keyboard Layout\Preload\1", "00000409", "REG_SZ"
    shell.RegWrite "HKEY_USERS\.DEFAULT\Keyboard Layout\Preload\2", "00000419", "REG_SZ"

    '2. Set keboard layout (English default)
    shell.RegWrite "HKEY_CURRENT_USER\Keyboard Layout\Preload\1",   "00000409", "REG_SZ"
    shell.RegWrite "HKEY_CURRENT_USER\Keyboard Layout\Preload\2",   "00000419", "REG_SZ"

    '3. Set hot key for change input language (Ctrl + Shift)
    shell.RegWrite "HKEY_CURRENT_USER\Keyboard Layout\Toggle\Hotkey",          2, "REG_SZ"
    shell.RegWrite "HKEY_CURRENT_USER\Keyboard Layout\Toggle\Language Hotkey", 2, "REG_SZ"
    shell.RegWrite "HKEY_CURRENT_USER\Keyboard Layout\Toggle\Layout Hotkey",   1, "REG_SZ"

    '4. Disable screen saver
    shell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveActive",   0, "REG_SZ"

    '5. Disable messages from notification area
    shell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\EnableBalloonTips",   0, "REG_DWORD"

    '6. Disable autostart CD
    shell.RegWrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\CDRom\Autorun",   0, "REG_DWORD"

    '7. Disable wizard which clear desktop shortcuts
    shell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\CleanupWiz\NoRun", 1, "REG_DWORD"
    WScript.Echo "Script is done!"
End Function