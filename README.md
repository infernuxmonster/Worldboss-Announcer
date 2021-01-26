# Worldboss Announcer

Simple PowerShell-script that works with the unitscan addon in order to post who is scouting what and when bosses spawn. This is the 2nd version, which incorperates an OCR-API in order to doublecheck that the text in the addon actually matches that of a boss. 

## Setup

1. The path `C:\Temp` needs to exist on your PC (or you can change the paths in the script).
2. Add your API-key from ocr.space on line 57 (`[string]$apiKey = ""`).
3. Add webhook URL from discord to line 121 and 228 (`$webHookUrl = ""`). Line 121 is for normal messages, 228 is for embedded notifications.
4. (Optional) - edit line 334 (`$Window.Icon = "https://cdn.discordapp.com/attachments/634238750964187147/727995811488596018/goose_2.jpg"`) to add your own icon to the UI.

## Known issues

This is a PoC PowerShell-code, as such it might have some issues or bugs.

### Script detects a worldboss without unitscan triggering

If you're running V1 the script might incorrectly trigger due to some colorchanges. The way the script works is that it checks the upper left corner for a combination of colors that are present in unitscan frames, but depending on the UI or colorscheme of the background it might trigger. 

This is fixed in version 2 with an OCR-API integration.

### Script doesn't work when tabbed out

That's right, it doesn't. The script only works when you have World of Warcraft running in fullscreen, as it only reads the screen, not the actual game itself.

### Script can't run due to ExecutionPolicy

To run script, you might have to set `ExecutionPolicy` to something else:

1. Open `powershell.exe` as administrator
2. Run `Set-ExecutionPolicy bypass`
3. Answer prompt with `Y` or `A`

## Help improve this simple script?

Just issue a pull request or dm me on twitter @infernuxmonster
