#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Date: 11/13/2019
# Description: Automates the installation of an application which doesn't have a
# silent switch by disabling user input and then injecting keystrokes to progress
# the installation.
#####
# Define C# method which imports Windows API exposing user BlockInput method
$methodDefinition = @"
    [DllImport("user32.dll")]
    public static extern bool BlockInput(bool fBlockIt);
"@

# Compile C# code, exposing it to PowerShell
$userInput = Add-Type -MemberDefinition $methodDefinition -Name UserInput -Namespace UserInput -PassThru

# Import Windows API
Add-Type -AssemblyName System.Windows.Forms

# Create an array of keystrokes in the order they are injected.
$keyStrokes = @("ENTER", "TAB", "ENTER", "TAB", "N", "TAB", "I", "ENTER")

# Create any strings that may need to be entered into dialogs.
# NOTE: The key should be the index of the keystroke the text will be entered prior to.
$textEntry = @{3 = "C:\Users\sathercd3383-lab\Desktop" }

# Define at which points we expect a long pause. This example is the second-to-last keystroke.
$pausePoints = @($keyStrokes.Count - 1)

# Since we cannot tell when the process has finished installing/running it's logic
# we need to define how long to wait during these periods, which is typically
# longer than what we would wait inbetween keystrokes. Set the key to be the index
# of the keystroke prior to the long pause.
# NOTE: The keys should match those in $pausePoints as this defines how long those pause points are.
$longPause = @{$($keyStrokes.Count - 1) = 35 }

# Wrap code in try-catch-finally so that if the script crashes, user input
# will be guaranteed to be re-enabled 
try {
    # Assign to desired dialog text
    $message = "Application installation will require user input to be disabled for ~45 seconds"

    # Display dialog, storing user selection
    $msgBoxInfo = [System.Windows.Forms.MessageBox]::Show($message, "University of Wisconsin - Eau Claire", "OkCancel", "Information")

    # If user selected 'OK'
    if ($msgBoxInfo -eq 1) {
        $userInput::BlockInput($true)

        # Place code to be executed while input blocked here
        # NOTE: Do NOT use -Wait as the script will not proceed until the process has finished.
        Start-Process -FilePath "C:\Users\sathercd3383-lab\Desktop\2019.6\2019.6_Part2\MetapluginWindows\installFull.exe"

        for ($i = 0; $i -lt $keyStrokes.Count; $i++) {
            if ($pausePoints.Contains($i)) {
                Start-Sleep -Seconds $longPause[$i]
            }
            else {
                Start-Sleep -Seconds 1
            }

            if ($textEntry[$i]) {
                foreach ($char in $textEntry[$i].ToCharArray()) {
                    [System.Windows.Forms.SendKeys]::SendWait("{$($char)}")
                }
            }

            [System.Windows.Forms.SendKeys]::SendWait("{$($keyStrokes[$i])}")
        }
    }
    else {
        exit
    }
}
catch {
    Write-Host "An error occurred:"
    Write-Host $_
}
finally {
    # Re-enable user input before exiting
    $userInput::BlockInput($false)
}