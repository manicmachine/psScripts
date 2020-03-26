#####
# Author: Corey Sather - sathercd3383@uwec.edu
# Description: Template for blocking user input while executing code. Ensures
# user input is restored prior to exiting.
#####
# Define C# method which imports Windows API exposing user BlockInput method
$methodDefinition = @"
    [DllImport("user32.dll")]
    public static extern bool BlockInput(bool fBlockIt);
"@

# Compile C# code, exposing it to PowerShell
$userInput = Add-Type -MemberDefinition $methodDefinition -Name UserInput -Namespace UserInput -PassThru

# Wrap code in try-catch-finally so that if the script crashes, user input
# will be guaranteed to be re-enabled 
try {
    $userInput::BlockInput($true) > $null

    # Place code to be executed while input blocked here
} catch {
    Write-Host "An error occurred:"
    Write-Host $_
} finally {
    # Re-enable user input before exiting
    $userInput::BlockInput($false) > $null
}