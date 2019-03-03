function Show-HtmlDialog {
    <#
        .SYNOPSIS
            For displaying an HTML formatted message in a dialog box.
    
        .DESCRIPTION
            Patterned off the dialogs in the PowerShell App Deployment Toolkit but uses WPF and HTML for scaling on high-DPI systems.
    
        .PARAMETER Message
            HTML to be displayed in the dialog. Mandatory.
    
        .PARAMETER Title
            What gets displayed in the window title.
    
        .PARAMETER Icon
            Path to the icon to be displayed in the title bar. Must be of file type ICO. Optional, will default to PowerShell icon.
    
        .PARAMETER Banner
            Path to the banner image to be displayed in the title bar. Must be of file type BMP, GIF, JPEG, JPG, PNG, TIFF, or TIF. Width of banner image will determine width of dialog. Optional.
    
        .PARAMETER SystemIcon
            System icon to display to the left of the message. Also chooses which system sound is played if the NoSound parameter is not used. Optional, if not included, the message box fills the dialog.

        .PARAMETER OkButton
            Shows an OK button in the lower right-hand corner of the dialog. When the OK button is clicked, the dialog will return "OK".
    
        .PARAMETER Timeout
            Time in seconds to display the dialog. When time runs out, the dialog will return "Timeout".
        
        .PARAMETER Progress
            Opens the dialog at the top of the screen and stays for the progress of a script. Requires closing at the end of the script.
    
        .PARAMETER NoSound
            Skip playing the system sound that corresponds to the SystemIcon parameter.
    
        .PARAMETER NoTimer
            Hides the countdown timer when used with the Timeout parameter.

        .PARAMETER NoClose
            Cancel closing the window when the X button in the upper right-hand corner of the window is clicked.
    
        .PARAMETER CloseProgress
            Closes the currently open progress dialog.
    
        .OUTPUTS
            [string]"OK" if the OK button is clicked
            [string]"Timeout" if the timeout is reached
            [string]"Progress" after the progress dialog is created
            [string]"CloseProgress" after the progress dialog is closed
            System.Boolean.False if the X in the title bar is clicked in a default dialog

        .EXAMPLE
            PS C:\>Show-HtmlDialog -Message "<p>Please remove the USB flash drive from the system.</p><br /><p>Click the X button to close this dialog.</p>"

            Shows dialog with the only mandatory parameter, Message. The icon in the title bar will be PowerShell's and the dialog will need to be closed with the X button in the title bar.
    
        .EXAMPLE
            PS C:\>Show-HtmlDialog -Title "Windows 10 Deployment" -Icon  ".\PreflightIcon.ico" -Banner  ".\PreflightBanner.png" -SystemIcon "Information" -Timeout 15 -Message "<p>Please remove the USB flash drive from the system.</p><br /><p>This dialog will timeout after 15 seconds and the Windows 10 deployment will continue.</p>"
            
            Shows the dialog with some extra, optional parameters specified, including Timeout.
    
        .EXAMPLE
            PS C:\>Show-HtmlDialog -Title "Windows 10 Deployment" -Icon  ".\PreflightIcon.ico" -Banner  ".\PreflightBanner.png" -SystemIcon "Error" -OkButton -Message "<p>You have not selected UEFI boot.</p><p>UEFI boot is required to deploy Windows 10.</p><p><b>Please restart the computer and enable UEFI boot before starting the deployment.</b></p>"
            
            Shows the dialog with some extra, optional parameters specified, including OkButton.

        .EXAMPLE
            PS C:\>Show-HtmlDialog -Progress -Message "<div align='center'><p>Downloading and applying the Windows 10 image.</p><p><i>Please wait...</i></p></div>"
            PS C:\>Show-HtmlDialog -CloseProgress

            Shows the dialog in a runspace, returning focus back to the calling script. The CloseProgress paramter is used to close the progress when desired.

        .EXAMPLE
            PS C:\>Show-HtmlProgress -Message "<div align='center'><p>Downloading and applying the Windows 10 image.</p><p><i>Please wait...</i></p></div>" -Icon  ".\PreflightIcon.ico" -Banner  ".\PreflightBanner.png" -SystemIcon "Information" -NoSound -NoClose
            PS C:\>Close-HtmlProgress

            Shows the function included to simplify access to the progress dialog.
            
        .LINK
            http://psappdeploytoolkit.com/
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ParameterSetName = "Default")]
        [Parameter(Mandatory=$true, ParameterSetName = "Progress")]
        [string]$Message, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [string]$Title = "PowerShell Dialog", 

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [ValidateScript({
            if (-not (Test-Path -Path $_)) {
                throw "File specified for the Icon argument does not exist"
            }
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "The Icon argument must be a file and not a folder path"
            }
            if ($_ -notmatch "(\.ico)") {
                throw "The file specified for the Icon argument must be of type ICO"
            }
            return $true
        })]
        [string]$Icon, 

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "The file specified for the Banner argument does not exist"
            }
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "The Banner argument must be a file and not a folder path"
            }
            if ($_ -notmatch "(\.bmp|\.gif|\.jpeg|\.jpg|\.png|\.tiff|\.tif)") {
                throw "The file specified for the Banner argument must be of type BMP, GIF, JPEG, JPG, PNG, TIFF, or TIF"
            }
            return $true
        })]
        [string]$Banner, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [ValidateSet("Application", "Asterisk", "Error", "Exclamation", "Hand", "Information", "Question", "Shield", "Warning", "WinLogo")]
        [string]$SystemIcon, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]$OkButton, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [int]$Timeout, 

        [Parameter(Mandatory = $true, ParameterSetName = "Progress")]
        [switch]$Progress, 

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [switch]$NoSound, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]$NoTimer,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [switch]$NoClose, 

        [Parameter(Mandatory = $true, ParameterSetName = "CloseProgress")]
        [switch]$CloseProgress
    )

    Add-Type -AssemblyName "PresentationFramework", "System.Drawing", "System.Windows.Forms", "WindowsFormsIntegration"

    $Return = $false

    # runs when the PowerShell console is closed
    if ($PsCmdlet.ParameterSetName -eq "Progress") {
        if (-not (Get-Event -SourceIdentifier PowerShell.Exiting -ErrorAction SilentlyContinue)) {
            Register-EngineEvent -SourceIdentifier PowerShell.Exiting -SupportEvent -Action {
                Show-HtmlDialog -CloseProgress
            }
        }
    }

    if ($PsCmdlet.ParameterSetName -ne "CloseProgress") {
        # creates or updates the dialog if $CloseProgress parameter NOT set
        
        if ($PsCmdlet.ParameterSetName -eq "Progress" -and $script:SynchronizedHashtable.XamlReader.Dispatcher.Thread.ThreadState -eq "Running") {
            # if the progress dialog is already running, update the message
            $script:SynchronizedHashtable.Message = $Message

            $script:SynchronizedHashtable.XamlReader.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]"Normal", 
                [System.Action]{ $script:SynchronizedHashtable.MessageBrowser.NavigateToString("
                    <!DOCTYPE html>
                    <html>
                        <body 
                            scroll='no' 
                            onselectstart='return false;' 
                            oncontextmenu='return false;' 
                            ondragover='window.event.returnValue=false;' 
                            style='cursor: default;'
                        >
                            $($script:SynchronizedHashtable.Message)
                        </body>
                    </html>
                ") }
            )
            
            $script:SynchronizedHashtable.XamlReader.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]"Normal", 
                [System.Action]{ $script:SynchronizedHashtable.XamlReader.Tag = "UpdateProgress" }
            )

            $Return = $script:SynchronizedHashtable.XamlReader.Tag #"UpdateProgress" #$true               
        }
        else {
            if ($script:SynchronizedHashtable.XamlReader.Dispatcher.Thread.ThreadState -eq "Running") {
                # if the progress dialog is already running, close it
                Show-HtmlDialog -CloseProgress
            }

            # if neither $OkButton and $Timeout parameter are used, set $NoClose to $false to allow way for closing dialog
            if ($NoClose -and -not ($OkButton -or $Timeout -or $Progress)) {
                $NoClose = $false
            }
            
            $IconPath = (Get-ChildItem -Path $Icon).FullName
            $BannerPath = (Get-ChildItem -Path $Banner).FullName
            
            $AppliedDPI = (Get-ItemProperty "HKCU:\Control Panel\Desktop\WindowMetrics" -Name "AppliedDPI").AppliedDPI
            $DpiPercentage = $AppliedDPI / 96

            $script:SynchronizedHashtable = [hashtable]::Synchronized(@{}) # for passing variables to and from runspace; script variable for closing progress dialog
            $script:SynchronizedHashtable.Message = $Message
            $script:SynchronizedHashtable.Title = $Title
            $script:SynchronizedHashtable.Icon = $Icon
            $script:SynchronizedHashtable.Banner = $Banner
            $script:SynchronizedHashtable.IconPath = $IconPath
            $script:SynchronizedHashtable.BannerPath = $BannerPath
            $script:SynchronizedHashtable.SystemIcon = $SystemIcon
            $script:SynchronizedHashtable.OkButton = $OkButton
            $script:SynchronizedHashtable.Timeout = $Timeout
            $script:SynchronizedHashtable.Progress = $Progress
            $script:SynchronizedHashtable.NoSound = $NoSound
            $script:SynchronizedHashtable.NoTimer = $NoTimer
            $script:SynchronizedHashtable.NoClose = $NoClose
            $script:SynchronizedHashtable.DpiPercentage = $DpiPercentage

            $script:Runspace = [runspacefactory]::CreateRunspace() # runspace used for progress dialog to allow synchronous scripting; script variable for closing runspace after progress dialog closes
            $script:SynchronizedHashtable.Runspace = $script:Runspace
            $script:Runspace.ApartmentState = "STA" #MTA, STA
            $script:Runspace.ThreadOptions = "ReuseThread" #Default, ReuseThread, UseCurrentThread, UseNewThread
            $script:Runspace.Open()
            $script:Runspace.SessionStateProxy.SetVariable("SynchronizedHashtable", $script:SynchronizedHashtable)
            
            $PowerShellCommand = [PowerShell]::Create().AddScript({
                $SystemIconDrawing = switch ($script:SynchronizedHashtable.SystemIcon) {
                    "Application" { [Drawing.SystemIcons]::Application }
                    "Asterisk" { [Drawing.SystemIcons]::Asterisk }
                    "Error" { [Drawing.SystemIcons]::Error }
                    "Exclamation" { [Drawing.SystemIcons]::Exclamation }
                    "Hand" { [Drawing.SystemIcons]::Hand }
                    "Information" { [Drawing.SystemIcons]::Information }
                    "Question" { [Drawing.SystemIcons]::Question }
                    "Shield" { [Drawing.SystemIcons]::Shield }
                    "Warning" { [Drawing.SystemIcons]::Warning }
                    "WinLogo" { [Drawing.SystemIcons]::WinLogo }
                }

                ######################################
                # set WebBrowser width offset depending upon if system icon choosen
                if ($SystemIconDrawing) {
                    $BrowserWidthOffset = 0
                    $BrowserPaddingOffset = 32 + 30 # icon width plus margin
                }
                else {
                    $BrowserWidthOffset = 32 + 30 # icon width plus margin
                    $BrowserPaddingOffset = 0
                }
                
                ######################################
                # build XamlReader form
                $XmlNodeReader = New-Object System.Xml.XmlNodeReader([xml]"<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'></Window>")
                $script:SynchronizedHashtable.XamlReader = [Windows.Markup.XamlReader]::Load($XmlNodeReader)
                
                $script:SynchronizedHashtable.XamlReader.Title = $script:SynchronizedHashtable.Title
                if ($script:SynchronizedHashtable.IconPath) {
                    <#-------------------------------------
                    unable to use BitMapImage directly or indirectly from WinPE x64 due to MSCMS.dll being only a 32-bit file
                    $script:SynchronizedHashtable.XamlReader.Icon = $script:SynchronizedHashtable.IconPath #>
                    $IconFileStream = New-Object System.IO.FileStream($script:SynchronizedHashtable.IconPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                    $IconBitmapDecoder = New-Object System.Windows.Media.Imaging.IconBitmapDecoder($IconFileStream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::None)
                    $script:SynchronizedHashtable.XamlReader.Icon = $IconBitmapDecoder.Frames[0]
                    #-------------------------------------
                }
                $script:SynchronizedHashtable.XamlReader.SizeToContent = "WidthAndHeight" #Manual (default), Height, Width, WidthAndHeight
                $script:SynchronizedHashtable.XamlReader.ResizeMode = "NoResize" #CanResize, NoResize, CanMinimize, CanResizeWithGrip
                $script:SynchronizedHashtable.XamlReader.Topmost = $true
                $script:SynchronizedHashtable.XamlReader.Tag = $false # WPF returns $true or $false with ShowDialog, so using the XamlReader's Tag property to return the DialogResult
                #$script:SynchronizedHashtable.XamlReader.ShowInTaskbar = $false

                $script:SynchronizedHashtable.XamlReader.WindowStartupLocation = "CenterScreen" #Manual, CenterOwner, CenterScreen                    
                if ($script:SynchronizedHashtable.Progress) {
                    $script:SynchronizedHashtable.XamlReader.WindowStartupLocation = "Manual" #Manual, CenterOwner, CenterScreen
                    $script:SynchronizedHashtable.XamlReader.Add_ContentRendered({    
                        $PrimaryScreen = [Windows.Forms.Screen]::PrimaryScreen
                        $PrimaryScreenWidth = $PrimaryScreen.WorkingArea.Width
                        $PrimaryScreenHeight = $PrimaryScreen.WorkingArea.Height
                        $script:SynchronizedHashtable.XamlReader.Left = (($PrimaryScreenWidth / 2) / $script:SynchronizedHashtable.DpiPercentage) - (($script:SynchronizedHashtable.XamlReader.Width) / 2)
                        $script:SynchronizedHashtable.XamlReader.Top = 30 #$PrimaryScreenHeight / 9.5
                        #[System.Windows.MessageBox]::Show($script:SynchronizedHashtable.XamlReader.Width)
                    })
                }

                if ($PsCmdlet.ParameterSetName -eq "Progress") {
                    $script:SynchronizedHashtable.XamlReader.Tag = "Progress"
                }

                ######################################
                # build grid and add rows
                $Grid = New-Object System.Windows.Controls.Grid
                #$Grid.ShowGridLines = $true
                $BannerRow = New-Object System.Windows.Controls.RowDefinition # Row 0
                $BannerRow.Height = "*" #[System.Windows.GridUnitType]::Star
                $Grid.RowDefinitions.Add($BannerRow)
                $MessageRow = New-Object System.Windows.Controls.RowDefinition # Row 1
                $MessageRow.Height = "4*"
                $Grid.RowDefinitions.Add($MessageRow)
                if ($script:SynchronizedHashtable.OkButton -or $script:SynchronizedHashtable.Timeout) {
                    $BottomRow = New-Object System.Windows.Controls.RowDefinition # Row 2
                    $BottomRow.Height = "60"
                    $Grid.RowDefinitions.Add($BottomRow)
                }
                
                ######################################
                # build banner and add to grid
                if ($script:SynchronizedHashtable.BannerPath) {
                    $BannerExtension = [IO.Path]::GetExtension($script:SynchronizedHashtable.BannerPath)
                    $BannerWidth = [System.Drawing.Image]::FromFile($script:SynchronizedHashtable.BannerPath).Width
                    $BannerHeight = [System.Drawing.Image]::FromFile($script:SynchronizedHashtable.BannerPath).Height
                    $BannerImage = New-Object System.Windows.Controls.Image
                    <#--------------------------------------
                    unable to use BitmapImage directly or indirectly from WinPE x64 due to MSCMS.dll being only a 32-bit file
                    #$BannerImage.Source = $script:SynchronizedHashtable.BannerPath
                    $BannerImage.Source = New-Object System.Windows.Media.Imaging.BitmapImage(New-Object uri($script:SynchronizedHashtable.BannerPath)) #>
                    $BannerFileStream = New-Object System.IO.FileStream($script:SynchronizedHashtable.BannerPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                    $BannerBitmapDecoder = switch -Regex ($BannerExtension) {
                        "BMP" { New-Object System.Windows.Media.Imaging.BmpBitmapDecoder($BannerFileStream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::None) }
                        "GIF" { New-Object System.Windows.Media.Imaging.GifBitmapDecoder($BannerFileStream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::None) }
                        "JPEG|JPG" { New-Object System.Windows.Media.Imaging.JpegBitmapDecoder($BannerFileStream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::None) }
                        "PNG" { New-Object System.Windows.Media.Imaging.PngBitmapDecoder($BannerFileStream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::None) }
                        "TIFF|TIF" { New-Object System.Windows.Media.Imaging.TiffBitmapDecoder($BannerFileStream, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::None) }
                    }
                    
                    $BannerImage.Source = $BannerBitmapDecoder.Frames[0]
                    #-------------------------------------
                    $BannerImage.Width = $BannerWidth #512
                    #$BannerImage.Stretch = [System.Windows.Media.Stretch]::Uniform
                    #$BannerImage.VerticalAlignment = "Top" #Stretch, Bottom, Center, Top
                    [System.Windows.Controls.Grid]::SetRow($BannerImage, 0) # BannerRow
                    [void]$Grid.Children.Add($BannerImage)
                }

                ######################################
                # build system icon and add to grid, if requested
                if ($SystemIconDrawing) {
                    $IconImage = New-Object System.Windows.Controls.Image
                    $IconImage.Source = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($SystemIconDrawing.Handle, [System.Windows.Int32Rect]::Empty, [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
                    $IconImage.Width = 32
                    $IconImage.HorizontalAlignment = "Left" #Stretch, Center, Left, Right
                    $IconImage.VerticalAlignment = "Top" #Stretch, Bottom, Center, Top
                    $IconImage.Margin = 15 #"15, 15, 15, 15" #"left, top, right, bottom"
                    [System.Windows.Controls.Grid]::SetRow($IconImage, 1) # MessageRow
                    [void]$Grid.Children.Add($IconImage)
                }

                ######################################
                # build web browser and add to grid
                $MessageBrowser = New-Object System.Windows.Controls.WebBrowser
                if ($BannerWidth) {
                    $MessageBrowser.Width = $BannerWidth - $BrowserPaddingOffset - 30
                }
                else {
                    $MessageBrowser.MinWidth = 414 + $BrowserWidthOffset
                    $MessageBrowser.Width = 0 #
                }
                $MessageBrowser.Height = 0 #224
                #$MessageBrowser.HorizontalAlignment = "Stretch" #Stretch, Center, Left, Right
                #$MessageBrowser.VerticalAlignment = "Top" #Stretch, Bottom, Center, Top
                $MessageBrowser.Margin = "$(15 + $BrowserPaddingOffset), 15, 15, 15" #"left, top, right, bottom"
                $MessageBrowser.NavigateToString("
                    <!DOCTYPE html>
                    <html>
                        <body 
                            scroll='no' 
                            onselectstart='return false;' 
                            oncontextmenu='return false;' 
                            ondragover='window.event.returnValue=false;' 
                            style='cursor: default;'
                        >
                            $($script:SynchronizedHashtable.Message)
                        </body>
                    </html>
                ")
                # resize browser control after it's loaded to fit everything in the window
                $MessageBrowser.Add_LoadCompleted({
                    $MessageBrowser.Width = $MessageBrowser.Document.body.parentElement.scrollWidth / $script:SynchronizedHashtable.DpiPercentage
                    $MessageBrowser.Height = $MessageBrowser.Document.body.parentElement.scrollHeight / $script:SynchronizedHashtable.DpiPercentage
                })
                $MessageBrowser.Name = "MessageBrowser"
                [System.Windows.Controls.Grid]::SetRow($MessageBrowser, 1) # MessageRow
                [void]$Grid.Children.Add($MessageBrowser)
                $script:SynchronizedHashtable.MessageBrowser = $MessageBrowser #$script:SynchronizedHashtable.XamlReader.FindName("MessageBrowser")
                
                ######################################
                # build Ok button and add to grid
                if ($script:SynchronizedHashtable.OkButton) {
                    $OkButtonControl = New-Object System.Windows.Controls.Button
                    $OkButtonControl.Content = "OK"
                    $OkButtonControl.Width = 100
                    $OkButtonControl.Height = 30
                    $OkButtonControl.HorizontalAlignment = "Right" #Stretch, Center, Left, Right
                    #$OkButtonControl.VerticalAlignment = "Bottom" #Stretch, Bottom, Center, Top
                    $OkButtonControl.Margin = 15 #"15, 15, 15, 15" #"left, top, right, bottom"
                    $OkButtonControl.IsDefault = $true
                    $OkButtonControl.Add_Click({
                        if ($Timer) {
                            $Timer.Stop()
                        }
                        $script:SynchronizedHashtable.XamlReader.Tag = "OK" #[System.Windows.Forms.DialogResult]::OK
                        #$script:SynchronizedHashtable.XamlReader.DialogResult = $script:SynchronizedHashtable.XamlReader.Tag # using Close() instead since now using ShowDialog()
                        $script:SynchronizedHashtable.XamlReader.Close()
                    })
                    [System.Windows.Controls.Grid]::SetRow($OkButtonControl, 2) # BottomRow
                    [void]$Grid.Children.Add($OkButtonControl)
                }

                ######################################
                # create timer
                if ($script:SynchronizedHashtable.Timeout) {
                    $Timer = New-Object -TypeName "System.Windows.Forms.Timer"
                    $Timer.Interval = 1000 # 1 second in milliseconds
                    $Timer.Tag = $script:SynchronizedHashtable.Timeout # used to keep track of time remaining

                    if (-not $script:SynchronizedHashtable.NoTimer) {
                        if ($script:SynchronizedHashtable.OkButton) {
                            $OkButtonControl.Content = "OK ($($Timer.Tag))"
                        }
                        else {
                            $TimerLabel = New-Object System.Windows.Controls.Label
                            #$TimerLabel.Width = 100
                            $TimerLabel.Height = 30
                            $TimerLabel.Margin = 15 #"15, 15, 15, 15" #"left, top, right, bottom"
                            $TimerLabel.HorizontalAlignment = "Right" #Stretch, Center, Left, Right
                            #$TimerLabel.VerticalAlignment = "Bottom" #Stretch, Bottom, Center, Top
                            $TimerLabel.HorizontalContentAlignment = "Right" #Stretch, Center, Left, Right
                            $TimerLabel.Content = "($($script:SynchronizedHashtable.Timeout))"
                            [System.Windows.Controls.Grid]::SetRow($TimerLabel, 2) # BottomRow
                            [void]$Grid.Children.Add($TimerLabel)
                        }
                    }

                    $Timer.Add_Tick({
                        $Timer.Tag -= 1
                        if ($Timer.Tag -eq 0) {
                            $script:SynchronizedHashtable.XamlReader.Tag = "Timeout" #[System.Windows.Forms.DialogResult]::None
                            #$script:SynchronizedHashtable.XamlReader.DialogResult = $script:SynchronizedHashtable.XamlReader.Tag # using Close() instead since now using ShowDialog()
                            $script:SynchronizedHashtable.XamlReader.Close()
                        } else {
                            if (-not $script:SynchronizedHashtable.NoTimer) {
                                if ($script:SynchronizedHashtable.OkButton) {
                                    $OkButtonControl.Content = "OK ($($Timer.Tag))"                    
                                }
                                else {
                                    $TimerLabel.Content = "($($Timer.Tag))"
                                }
                            }
                        }
                    })
                    $Timer.Start()
                }

                ######################################
                # add grid to form
                $script:SynchronizedHashtable.XamlReader.Content = $Grid
                
                    ######################################
                # play sound depending upon icon
                if (-not $script:SynchronizedHashtable.NoSound) {
                    switch ($script:SynchronizedHashtable.SystemIcon) {
                        "Application" { [System.Media.SystemSounds]::Beep.Play() } #Asterisk, Beep, Exclamation, Hand, Question
                        "Asterisk" { [System.Media.SystemSounds]::Asterisk.Play() }
                        "Error" { [System.Media.SystemSounds]::Hand.Play() }
                        "Exclamation" { [System.Media.SystemSounds]::Exclamation.Play() }
                        "Hand" { [System.Media.SystemSounds]::Hand.Play() }
                        "Information" { [System.Media.SystemSounds]::Beep.Play() }
                        "Question" { [System.Media.SystemSounds]::Beep.Play() }
                        "Shield" { [System.Media.SystemSounds]::Beep.Play() }
                        "Warning" { [System.Media.SystemSounds]::Exclamation.Play() }
                        "WinLogo" { [System.Media.SystemSounds]::Beep.Play() }
                    }
                }

                $script:SynchronizedHashtable.XamlReader.Add_Closing({
                    # cancel close if $NoClose parameter used
                    if ($script:SynchronizedHashtable.NoClose -and $script:SynchronizedHashtable.XamlReader.Tag -notin "OK", "Timeout", "CloseProgress") {
                        $_.Cancel = $true
                    }
                    else {
                        if ($Timer) {
                            $Timer.Stop()
                        }
                        
                        # clean up from using file stream
                        if ($IconFileStream) {
                            $IconFileStream.Close()
                            $IconFileStream.Dispose()                    
                        }
                        if ($BannerFileStream) {
                            $BannerFileStream.Close()
                            $BannerFileStream.Dispose()
                        }

                        [System.Windows.Forms.Application]::Exit() # for use with $ApplicationContext
                        #Stop-Process $PID
                    }
                })

                ######################################
                # Show form
                <#--------------------------------------
                instead of ShowDialog, using the code below from https://blog.netnerds.net/2016/01/showdialog-sucks-use-applicationcontexts-instead/
                $script:SynchronizedHashtable.XamlReader.ShowDialog() | Out-Null #>
                #[System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($script:SynchronizedHashtable.XamlReader) # allow input to window for TextBoxes, etc
                $script:SynchronizedHashtable.XamlReader.Show() # running this without $ApplicationContext and ::Run would actually cause a really poor response
                $script:SynchronizedHashtable.XamlReader.Activate() | Out-Null # this makes it pop up
                $ApplicationContext = New-Object "System.Windows.Forms.ApplicationContext" # create an application context for it to all run within; this helps with responsiveness and threading
                [System.Windows.Forms.Application]::Run($ApplicationContext) | Out-Null
                #-------------------------------------

                return $script:SynchronizedHashtable.XamlReader.Tag
            })
            $PowerShellCommand.Runspace = $script:Runspace
            if ($PsCmdlet.ParameterSetName -eq "Progress") {
                # if $Progress parameter, don't provide input and ouput variable to PowerShell instance so script continues
                $PowerShellAsyncResult = $PowerShellCommand.BeginInvoke()
                $Return = $script:SynchronizedHashtable.XamlReader.Tag #"Progress" #$true
            }
            elseif ($PsCmdlet.ParameterSetName -ne "CloseProgress") {
                # if neither $Progress nor $CloseProgress parameter, provide input and ouput variable to PowerShell instance to get output of dialog
                $PSDataCollection = New-Object "System.Management.Automation.PSDataCollection[psobject]"
                $PowerShellAsyncResult = $PowerShellCommand.BeginInvoke($PSDataCollection, $PSDataCollection)
                $Return = $PSDataCollection
                <#--------------------------------------
                this is another way to wait for output from the dialog
                do { 
                    Start-Sleep -Milliseconds 100
                } while (-not $PowerShellAsyncResult.IsCompleted)
                $PowerShellCommand.EndInvoke($PowerShellAsyncResult)
                $PowerShellCommand.Dispose()
                $script:Runspace.Close()
                $Return = $script:SynchronizedHashtable.XamlReader.Tag
                --------------------------------------#>
            }
        }
    }
    else {
        # runs if $CloseProgress parameter set
        if ($script:SynchronizedHashtable.XamlReader.Dispatcher.Thread.ThreadState -eq "Running") {
            # set XamlReader Tag property to CloseProgress so the dialog will close
            $script:SynchronizedHashtable.XamlReader.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]"Normal", 
                [System.Action]{ $script:SynchronizedHashtable.XamlReader.Tag = "CloseProgress" }
            )
            # close the XamlReader window
            $script:SynchronizedHashtable.XamlReader.Dispatcher.Invoke(
                [System.Windows.Threading.DispatcherPriority]"Normal", 
                [System.Action]{ $script:SynchronizedHashtable.XamlReader.Close() }
            )
            $script:SynchronizedHashtable.Clear()
            $script:Runspace.Close()

            Unregister-Event -SourceIdentifier PowerShell.Exiting -ErrorAction SilentlyContinue
        }

        $Return = $script:SynchronizedHashtable.XamlReader.Tag #"CloseProgress" #$true
    }

    return $Return
}



function Show-HtmlProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ParameterSetName = "Progress")]
        [string]$Message, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [string]$Title = "PowerShell Progress", 

        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [ValidateScript({
            if (-not (Test-Path -Path $_)) {
                throw "File specified for the Icon argument does not exist"
            }
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "The Icon argument must be a file and not a folder path"
            }
            if ($_ -notmatch "(\.ico)") {
                throw "The file specified for the Icon argument must be of type ICO"
            }
            return $true
        })]
        [string]$Icon, 

        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "The file specified for the Banner argument does not exist"
            }
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "The Banner argument must be a file and not a folder path"
            }
            if ($_ -notmatch "(\.bmp|\.gif|\.jpeg|\.jpg|\.png|\.tiff|\.tif)") {
                throw "The file specified for the Banner argument must be of type BMP, GIF, JPEG, JPG, PNG, TIFF, or TIF"
            }
            return $true
        })]
        [string]$Banner, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [ValidateSet("Application", "Asterisk", "Error", "Exclamation", "Hand", "Information", "Question", "Shield", "Warning", "WinLogo")]
        [string]$SystemIcon, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [switch]$NoSound, 
        
        [Parameter(Mandatory = $false, ParameterSetName = "Progress")]
        [switch]$NoClose, 

        [Parameter(Mandatory = $true, ParameterSetName = "CloseProgress")]
        [switch]$CloseProgress
    )

    if ($PsCmdlet.ParameterSetName -eq "CloseProgress") {
        Show-HtmlDialog -CloseProgress
    }
    else {
        $Splat = @{
            "Progress" = $true
            "Message" = $Message
            "Title" = $Title
        }
        if ($Icon) { $Splat.Add("Icon", $Icon) }
        if ($Icon) { $Splat.Add("Banner", $Banner) }
        if ($Icon) { $Splat.Add("SystemIcon", $SystemIcon) }
        if ($Icon) { $Splat.Add("NoSound", $NoSound) }
        if ($Icon) { $Splat.Add("NoClose", $NoClose) }
        
        Show-HtmlDialog @Splat    
    }
}



function Close-HtmlProgress {
    Show-HtmlDialog -CloseProgress
}
