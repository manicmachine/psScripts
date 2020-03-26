#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Description: Template for easily creating Windows shortcuts dynamically
#####
# Path to application's executable.
$targetFile = "" 

# Add argumenets if there are any
$arguments = ""

# Path which the shortcut will be created; change as needed
$shortcutFile = "$env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\"

####
# Don't touch this section.
####
$wScriptShell = New-Object -ComObject WScript.Shell
$shortcut = $wScriptShell.CreateShortcut($shortcutFile)
$shortcut.TargetPath = $targetFile
$shortcut.Arguments = $arguments
$shortcut.Save()