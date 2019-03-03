# Show-HtmlDialog

For displaying an HTML formatted message in a dialog box.

## Description

Patterned off the dialogs in the PowerShell App Deployment Toolkit but uses WPF and HTML for scaling on high-DPI systems.

Also included in this script are the Show-HtmlProgress and Close-HtmlProgress functions, added to make it easier to interface with the Progress parameter of the Show-HtmlDialog function.

## History

I was using the PSAppDeployToolkit to show notifications with my Windows 10 deployment but the font was real small on high-DPI displays. To fix this, I thought it would be good to try my hand at creating my own dialog. I created it so I could also use HTML in the message to provide some formatting.

I'm including some of the help from the script below.

## .PARAMETER Message
    HTML to be displayed in the dialog. Mandatory.

## .PARAMETER Title
    What gets displayed in the window title.

## .PARAMETER Icon
    Path to the icon to be displayed in the title bar. Must be of file type ICO. Optional, will default to PowerShell icon.

## .PARAMETER Banner
    Path to the banner image to be displayed in the title bar. Must be of file type BMP, GIF, JPEG, JPG, PNG, TIFF, or TIF. Width of banner image will determine width of dialog. Optional.

## .PARAMETER SystemIcon
    System icon to display to the left of the message. Also chooses which system sound is played if the NoSound parameter is not used. Optional, if not included, the message box fills the dialog.

## .PARAMETER OkButton
    Shows an OK button in the lower right-hand corner of the dialog. When the OK button is clicked, the dialog will return "OK".

## .PARAMETER Timeout
    Time in seconds to display the dialog. When time runs out, the dialog will return "Timeout".

## .PARAMETER Progress
    Opens the dialog at the top of the screen and stays for the progress of a script. Requires closing at the end of the script.

## .PARAMETER NoSound
    Skip playing the system sound that corresponds to the SystemIcon parameter.

## .PARAMETER NoTimer
    Hides the countdown timer when used with the Timeout parameter.

## .PARAMETER NoClose
    Cancel closing the window when the X button in the upper right-hand corner of the window is clicked.

## .PARAMETER CloseProgress
    Closes the currently open progress dialog.

## .OUTPUTS
    [string]"OK" if the OK button is clicked
    [string]"Timeout" if the timeout is reached
    [string]"Progress" after the progress dialog is created
    [string]"CloseProgress" after the progress dialog is closed
    System.Boolean.False if the X in the title bar is clicked in a default dialog

## .EXAMPLE
    PS C:\>Show-HtmlDialog -Message "<p>Please remove the USB flash drive from the system.</p><br /><p>Click the X button to close this dialog.</p>"

    Shows dialog with the only mandatory parameter, Message. The icon in the title bar will be PowerShell's and the dialog will need to be closed with the X button in the title bar.

## .EXAMPLE
    PS C:\>Show-HtmlDialog -Title "Windows 10 Deployment" -Icon  ".\PreflightIcon.ico" -Banner  ".\PreflightBanner.png" -SystemIcon "Information" -Timeout 15 -Message "<p>Please remove the USB flash drive from the system.</p><br /><p>This dialog will timeout after 15 seconds and the Windows 10 deployment will continue.</p>"
    
    Shows the dialog with some extra, optional parameters specified, including Timeout.

## .EXAMPLE
    PS C:\>Show-HtmlDialog -Title "Windows 10 Deployment" -Icon  ".\PreflightIcon.ico" -Banner  ".\PreflightBanner.png" -SystemIcon "Error" -OkButton -Message "<p>You have not selected UEFI boot.</p><p>UEFI boot is required to deploy Windows 10.</p><p><b>Please restart the computer and enable UEFI boot before starting the deployment.</b></p>"
    
    Shows the dialog with some extra, optional parameters specified, including OkButton.

## .EXAMPLE
    PS C:\>Show-HtmlDialog -Progress -Message "<div align='center'><p>Downloading and applying the Windows 10 image.</p><p><i>Please wait...</i></p></div>"
    PS C:\>Show-HtmlDialog -CloseProgress

    Shows the dialog in a runspace, returning focus back to the calling script. The CloseProgress paramter is used to close the progress when desired.

## .EXAMPLE
    PS C:\>Show-HtmlProgress -Message "<div align='center'><p>Downloading and applying the Windows 10 image.</p><p><i>Please wait...</i></p></div>" -Icon  ".\PreflightIcon.ico" -Banner  ".\PreflightBanner.png" -SystemIcon "Information" -NoSound -NoClose
    PS C:\>Close-HtmlProgress

    Shows the function included to simplify access to the progress dialog.
    
## .LINK
    http://psappdeploytoolkit.com/