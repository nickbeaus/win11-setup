# Windows Configuration Script - Interactive PowerShell Version
# Automatically runs as Administrator

# Check if running as Administrator
Clear-Host
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires Administrator privileges. Relaunching as Administrator..." -ForegroundColor Yellow
    
    # Get the script path
    $scriptPath = $MyInvocation.MyCommand.Path
    
    # Relaunch the script with Administrator privileges
    Start-Process PowerShell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    
    # Exit the current non-elevated process
    exit
}

Clear-Host
# Resize the console window
$Host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(100, 50)

function Header {
    Write-Host "Windows Configuration Script" -ForegroundColor Cyan
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "Running with Administrator privileges" -ForegroundColor Green
    Write-Host "==============================" -ForegroundColor Cyan
}

# Display the header
Header

function Prompt-User {
    param([string]$Message)
    do {
        $response = Read-Host "$Message (y/n)"
        $response = $response.ToLower()
        if ($response -eq '') { $response = 'y' }  # Default to 'y' if user just presses Enter
    } while ($response -ne 'y' -and $response -ne 'n')
    return $response -eq 'y'
}

# ask if the user wants to change timezone

# Ask if user wants to change timezone
Write-Host "`n"
if (Prompt-User "Would you like to set your timezone?") {
    Clear-Host
    Header
    Write-Host "Setting Timezone`n" -ForegroundColor Cyan
    # Mark that the user wants to set the timezone
    $setTimezone = $true
    # Get major US timezones only
    $majorTimezones = @(
        "Eastern Standard Time",
        "Central Standard Time", 
        "Mountain Standard Time",
        "Pacific Standard Time",
        "Alaska Standard Time",
        "Hawaiian Standard Time"
    )

    $timezones = Get-TimeZone -ListAvailable | Where-Object { $_.Id -in $majorTimezones } | Sort-Object DisplayName

    # Display timezones
    Write-Host "Available timezones:"
    for ($i = 0; $i -lt $timezones.Count; $i++) {
        Write-Host "$($i + 1). $($timezones[$i].DisplayName)"
    }
    Write-Host "`n"
    # Get user selection
    do {
        $selection = Read-Host "Enter the number of your timezone (1-$($timezones.Count))"
        $selectionNum = 0
        $validSelection = [int]::TryParse($selection, [ref]$selectionNum) -and 
                         $selectionNum -ge 1 -and 
                         $selectionNum -le $timezones.Count
    } while (!$validSelection)

    # Set the timezone
    $selectedTimezone = $timezones[$selectionNum - 1]
}

# Ask if the user wants to apply Windows Optimizations
Clear-Host
Header
Write-Host "`n"
if (Prompt-User "Would you like to apply Windows optimizations?") {
    Clear-Host
    Header
    Write-Host "Applying Windows Optimizations`n" -ForegroundColor Cyan

    # Mark that user wants to apply optimizations
    $WindowsOptimizations = $true
    # Enable compact mode for Windows Explorer
    if (Prompt-User "Enable compact mode for Windows Explorer?") {
        $Opt_Compact = $true
    } else {
        $Opt_Compact = $false
    }
    # Show file extensions
    if (Prompt-User "Show file extensions in Windows Explorer?") {
        $Opt_ShowFileExtensions = $true
    } else {
        $Opt_ShowFileExtensions = $false
    }
    # Show hidden files
    if (Prompt-User "Show hidden files in Windows Explorer?") {
        $Opt_ShowHiddenFiles = $true
    } else {
        $Opt_ShowHiddenFiles = $false
    }
    # Open File Explorer to This PC
    if (Prompt-User "Open File Explorer to This PC instead of Quick Access?") {
        $Opt_OpenThisPC = $true
    } else {
        $Opt_OpenThisPC = $false
    }
    # Show full path in title bar
    if (Prompt-User "Show full path in the title bar of File Explorer?") {
        $Opt_ShowFullPath = $true
    } else {
        $Opt_ShowFullPath = $false
    }
    # Enable verbose status messages
    if (Prompt-User "Enable verbose status messages during startup and shutdown?") {
        $Opt_Verbose = $true
    } else {
        $Opt_Verbose = $false
    }
    # Left align taskbar
    if (Prompt-User "Left align the taskbar icons?") {
        $Opt_LeftAlignTaskbar = $true
    } else {
        $Opt_LeftAlignTaskbar = $false
    }
    # Disable Bing in Start Menu
    if (Prompt-User "Disable Bing integration in the Start Menu?") {
        $Opt_DisableBing = $true
    } else {
        $Opt_DisableBing = $false
    }
    # Remove Search Bar
    if (Prompt-User "Remove the Search Bar from the Taskbar?") {
        $Opt_RemoveSearchBar = $true
    } else {
        $Opt_RemoveSearchBar = $false
    }
    # Remove chat icon
    if (Prompt-User "Remove the Chat icon from the Taskbar?") {
        $Opt_RemoveChatIcon = $true
    } else {
        $Opt_RemoveChatIcon = $false
    }
    # Disable shake minimize
    if (Prompt-User "Disable the 'Shake to Minimize' feature?") {
        $Opt_DisableShakeMinimize = $true
    } else {
        $Opt_DisableShakeMinimize = $false
    }
    # Disable greeting on first login
    if (Prompt-User "Disable the greeting on first login?") {
        $Opt_DisableGreeting = $true
    } else {
        $Opt_DisableGreeting = $false
    }
    # Default to "Do this for all current items"
    if (Prompt-User "Default to 'Do this for all current items' in file operations?") {
        $Opt_DefaultForAll = $true
    } else {
        $Opt_DefaultForAll = $false
    }
    # Add copy path to context menu
    if (Prompt-User "Add 'Copy Path' to the context menu?") {
        $Opt_CopyPath = $true
    } else {
        $Opt_CopyPath = $false
    }
}
# If user does not want to apply optimizations, set all options to false
else {
    $WindowsOptimizations = $false
    $Opt_Compact = $false
    $Opt_ShowFileExtensions = $false
    $Opt_ShowHiddenFiles = $false
    $Opt_OpenThisPC = $false
    $Opt_ShowFullPath = $false
    $Opt_Verbose = $false
    $Opt_LeftAlignTaskbar = $false
    $Opt_DisableBing = $false
    $Opt_RemoveSearchBar = $false
    $Opt_RemoveChatIcon = $false
    $Opt_DisableShakeMinimize = $false
    $Opt_DisableGreeting = $false
    $Opt_DefaultForAll = $false
    $Opt_CopyPath = $false
}

# Ask the user if they want to remove windows default Appx packages
Clear-Host
Header
Write-Host "`n"
if (Prompt-User "Do you want to remove default windows applications?") {
    Clear-Host
    Header
    Write-Host "Removing Default Windows Applications`n" -ForegroundColor Cyan
    # Mark that user wants to remove default Appx packages
    $RemoveDefaultAppx = $true
    $appxPackages = @(
        @{PackageId = "*clipchamp.clipchamp*"; Name = "Clipchamp"},
        @{PackageId = "*microsoft.copilot*"; Name = "Copilot"},
        @{PackageId = "*microsoft.windowscommunicationsapps*"; Name = "Mail and Calendar"},
        @{PackageId = "*microsoft.windowsfeedbackhub*"; Name = "Feedback Hub"},
        @{PackageId = "*microsoft.alarms*"; Name = "Alarms & Clock"},
        @{PackageId = "*microsoft.549981c3f5f10*"; Name = "Cortana"},
        @{PackageId = "*microsoft.windowsmaps*"; Name = "Maps"},
        @{PackageId = "*microsoft.bingnews*"; Name = "News"},
        @{PackageId = "*microsoft.todos*"; Name = "Microsoft To Do"},
        @{PackageId = "*microsoft.zunevideo*"; Name = "Movies & TV"},
        @{PackageId = "*microsoftcorporationii.quickassist*"; Name = "Quick Assist"},
        @{PackageId = "*microsoft.microsoftsolitairecollection*"; Name = "Microsoft Solitaire Collection"},
        @{PackageId = "*microsoft.microsoftstickynotes*"; Name = "Sticky Notes"},
        @{PackageId = "*microsoft.bingweather*"; Name = "Weather"},
        @{PackageId = "*microsoft.zunemusic*"; Name = "Groove Music"},
        @{PackageId = "*microsoft.windowssoundrecorder*"; Name = "Sound Recorder"},
        @{PackageId = "*microsoft.yourphone*"; Name = "Your Phone"},
        @{PackageId = "*microsoft.getstarted*"; Name = "Get Started"},
        @{PackageId = "*microsoft.outlook*"; Name = "Outlook"},
        @{PackageId = "*microsoft.gethelp*"; Name = "Get Help"},
        @{PackageId = "*microsoft.xboxapp*"; Name = "Xbox App"},
        @{PackageId = "*microsoft.gamingapp*"; Name = "Xbox Game Bar"},
        @{PackageId = "*microsoft.xboxidentityprovider*"; Name = "Xbox Identity Provider"},
        @{PackageId = "*microsoft.xboxspeechtotextoverlay*"; Name = "Xbox Speech to Text Overlay"},
        @{PackageId = "*microsoft.xboxgameoverlay*"; Name = "Xbox Game Overlay"},
        @{PackageId = "*microsoft.xboxgamingoverlay*"; Name = "Xbox Gaming Overlay"},
        @{PackageId = "*microsoft.xbox.cloudgaming*"; Name = "Xbox Cloud Gaming"},
        @{PackageId = "*microsoft.xbox.tcui*"; Name = "Xbox TCU"}
    )

    Write-Host "Please select the applications you want to remove:"
    Write-Host "Enter the numbers separated by commas (e.g., 1,5,10) or spaces (e.g., 1 5 10)."
    Write-Host "To remove all applications, type 'all'." -ForegroundColor Yellow
    Write-Host "-----------------------------------------"
    for ($i = 0; $i -lt $appxPackages.Count; $i++) {
        Write-Host "$($i + 1). $($appxPackages[$i].Name)"
    }
    Write-Host "-----------------------------------------"
    $selection = Read-Host "Enter your choice(s)"
        if ($selection -eq "all") {
        $selectedAppxObjects = $appxPackages
    } else {
        $inputTokens = $selection.Split([char[]](',', ' '), [System.StringSplitOptions]::RemoveEmptyEntries)
        $selectedAppxObjects = $inputTokens | ForEach-Object {
            [int]$num = 0 # Initialize $num for TryParse for each iteration
            if ([int]::TryParse($_, [ref]$num) -and $num -ge 1 -and $num -le $appxPackages.Length) {
                $appxPackages[$num - 1] # Output the object to the pipeline for collection
            } else {
                Write-Warning "Invalid selection: '$_' will be ignored."
                $null # Output $null for invalid entries so they can be filtered
            }
        } | Where-Object { $_ -ne $null } # Filter out any $null values
    }
    # Sort collected objects for consistent confirmation output (and remove duplicates if any)
    $selectedAppxObjects = $selectedAppxObjects | Sort-Object -Property DisplayName
}

# OneDrive Removal
# Ask if the user wants to remove OneDrive
Clear-Host
Header
Write-Host "`n"
if (Prompt-User "Do you want to remove OneDrive?") {
    $OneDriveRemoval = $true
}
else {
    $OneDriveRemoval = $false
}

Clear-Host
Header
Write-Host "`n"

if (Prompt-User "Do you want to install software?") {
    Clear-Host
    Header
    Write-Host "Installing Software`n" -ForegroundColor Cyan
    # Mark that user wants to install software
    $InstallSoftware = $true
    # Define the list of applications to install using Winget
        $applications = @(
        @{ PackageId = "7zip.7zip"; DisplayName = "7-Zip Archiver" },
        @{ PackageId = "angryziber.AngryIPScanner"; DisplayName = "Angry IP Scanner" },
        @{ PackageId = "Brave.Brave"; DisplayName = "Brave Browser" },
        @{ PackageId = "DBeaver.DBeaver.Community"; DisplayName = "DBeaver Community" },
        @{ PackageId = "Discord.Discord"; DisplayName = "Discord" },
        @{ PackageId = "Google.Chrome"; DisplayName = "Google Chrome" },
        @{ PackageId = "Google.EarthPro"; DisplayName = "Google Earth Pro" },
        @{ PackageId = "HandBrake.HandBrake"; DisplayName = "HandBrake" },
        @{ PackageId = "DuongDieuPhap.ImageGlass"; DisplayName = "ImageGlass Viewer" },
        @{ PackageId = "LIGHTNINGUK.ImgBurn"; DisplayName = "ImgBurn" },
        @{ PackageId = "apple.itunes"; DisplayName = "Apple iTunes" },
        @{ PackageId = "Mozilla.Firefox"; DisplayName = "Mozilla Firefox" },
        @{ PackageId = "Microsoft.Office"; DisplayName = "Microsoft Office 365" },
        @{ PackageId = "NDI.NDITools"; DisplayName = "NDI Tools" },
        @{ PackageID = "Notepad++.Notepad++"; DisplayName = "Notepad++" },
        @{ PackageId = "OBSProject.OBSStudio"; DisplayName = "OBS Studio" },
        @{ PackageId = "Prusa3D.PrusaSlicer"; DisplayName = "PrusaSlicer" },
        @{ PackageId = "PuTTY.PuTTY"; DisplayName = "PuTTY SSH Client" },
        @{ PackageId = "Valve.Steam"; DisplayName = "Steam" },
        @{ PackageId = "TeraTermProject.teraterm"; DisplayName = "TeraTerm" },
        @{ PackageId = "VideoLAN.VLC"; DisplayName = "VLC Media Player" },
        @{ PackageId = "WinSCP.WinSCP"; DisplayName = "WinSCP" },
        @{ PackageId = "WireGuard.WireGuard"; DisplayName = "WireGuard" },
        @{ PackageId = "WiresharkFoundation.Wireshark"; DisplayName = "Wireshark" },
        @{ PackageId = "AntibodySoftware.WizTree"; DisplayName = "WizTree Disk Analyzer" },
        @{ PackageId = "XnSoft.XnConvert"; DisplayName = "XnConvert" },
        @{ PackageId = "XnSoft.XnViewMP"; DisplayName = "XnView MP" },
        @{ PackageId = "Yubico.YubikeyManager"; DisplayName = "YubiKey Manager" },
        @{ PackageId = "Zoom.Zoom"; DisplayName = "Zoom" }
    )
    ## Script Logic

    Write-Host "Please select the software you want to install."
    Write-Host "Enter the numbers separated by commas (e.g., 1,5,10) or spaces (e.g., 1 5 10)."
    Write-Host "-----------------------------------------"

        # Display the list of applications to the user using their DisplayName
    for ($i = 0; $i -lt $applications.Length; $i++) {
        Write-Host "$($i + 1). $($applications[$i].DisplayName)"
    }

    Write-Host "-----------------------------------------"
    $selection = Read-Host "Enter your choice(s)"

    # --- Process User Input ---
    # This section parses the user's input and collects the selected applications.
    $inputTokens = $selection.Split([char[]](',', ' '), [System.StringSplitOptions]::RemoveEmptyEntries)

    # Collect valid selections directly from the pipeline
    $selectedAppObjects = $inputTokens | ForEach-Object {
        [int]$num = 0 # Initialize $num for TryParse for each iteration
        if ([int]::TryParse($_, [ref]$num) -and $num -ge 1 -and $num -le $applications.Length) {
            $applications[$num - 1] # Output the object to the pipeline for collection
        } else {
            Write-Warning "Invalid selection: '$_' will be ignored."
            $null # Output $null for invalid entries so they can be filtered
        }
    } | Where-Object { $_ -ne $null } # Filter out any $null values from invalid entries

    # Sort collected objects for consistent confirmation output (and remove duplicates if any)
    $selectedAppObjects = $selectedAppObjects | Sort-Object -Property DisplayName # Removed -Unique as it was filtering valid items
}

Clear-Host
Header
Write-Host "`n"
if (Prompt-User "Do you want to activate products?") {
    # Mark that user wants to activate products
    $activateProducts = $true
    Clear-Host
    Header
    Write-Host "Activating Products`n" -ForegroundColor Cyan
    if (Prompt-User "Do you want to activate Windows?") {
        $activateWindows = $true
    } else {
        $activateWindows = $false
    }
    if (Prompt-User "Do you want to activate Office?") {
        $activateOffice = $true
    } else {
        $activateOffice = $false
    }
}


# sleep 1 second
Start-Sleep -Seconds 1
Write-Host "`n-----------------------------------------" -ForegroundColor Cyan
Write-Host "Applying your selections..." -ForegroundColor Cyan
Write-Host "-----------------------------------------`n" -ForegroundColor Cyan
if ($setTimezone) {
    Write-Host "You have chosen to set the timezone." -ForegroundColor Green
    if ($selectedTimezone) {
    Write-Host "Setting timezone to: $($selectedTimezone.DisplayName)" -ForegroundColor Green
    Set-TimeZone -Id $selectedTimezone.Id
    } else {
        Write-Host "No timezone change selected." -ForegroundColor Yellow
    }
} else {
    Write-Host "You have chosen not to set the timezone." -ForegroundColor Yellow
}



# User wants to apply Windows optimizations
if ($WindowsOptimizations) {
    Write-Host "Applying Windows optimizations..." -ForegroundColor Green
    if ($Opt_Compact) {
        Write-Host "Enabling compact mode for Windows Explorer..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "UseCompactMode" -Value 1
        Write-Host "Compact mode enabled." -ForegroundColor Green
    } else {
        Write-Host "Compact mode not enabled." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Show file extensions
    if ($Opt_ShowFileExtensions) {
        Write-Host "Showing file extensions in Windows Explorer..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "HideFileExt" -Value 0
        Write-Host "File extensions are now visible." -ForegroundColor Green
    } else {
        Write-Host "File extensions will remain hidden." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Show hidden files
    if ($Opt_ShowHiddenFiles) {
        Write-Host "Showing hidden files in Windows Explorer..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "Hidden" -Value 1
        # Also set the "ShowSuperHidden" property to show system files
        Set-ItemProperty -Path $regPath -Name "ShowSuperHidden" -Value 1
        Write-Host "Hidden files are now visible." -ForegroundColor Green
    } else {
        Write-Host "Hidden files will remain hidden." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Open File Explorer to This PC
    if ($Opt_OpenThisPC) {
        Write-Host "Setting File Explorer to open to This PC..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "LaunchTo" -Value 1
        Write-Host "File Explorer will now open to This PC." -ForegroundColor Green
    } else {
        Write-Host "File Explorer will continue to open to Quick Access." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Show full path in title bar
    if ($Opt_ShowFullPath) {
        Write-Host "Showing full path in the title bar of File Explorer..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "FullPathAddress" -Value 1
        Write-Host "Full path will now be displayed in the title bar." -ForegroundColor Green
    } else {
        Write-Host "Full path will not be displayed in the title bar." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Enable verbose status messages
    if ($Opt_Verbose) {
        Write-Host "Enabling verbose status messages during startup and shutdown..." -ForegroundColor Green
        $regPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Set-ItemProperty -Path $regPath -Name "verbosestatus" -Value 1
        Write-Host "Verbose status messages enabled." -ForegroundColor Green
    } else {
        Write-Host "Verbose status messages will remain disabled." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Left align taskbar
    if ($Opt_LeftAlignTaskbar) {
        Write-Host "Left aligning the taskbar icons..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "TaskbarAl" -Value 0
        Write-Host "Taskbar icons are now left aligned." -ForegroundColor Green
    } else {
        Write-Host "Taskbar icons will remain centered." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Disable Bing in Start Menu
    if ($Opt_DisableBing) {
        Write-Host "Disabling Bing integration in the Start Menu..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
        New-Item -Path $regPath -Force | Out-Null
        Set-ItemProperty -Path $regPath -Name "DisableSearchBoxSuggestions" -Value 1
        Write-Host "Bing integration in the Start Menu has been disabled." -ForegroundColor Green
    } else {
        Write-Host "Bing integration in the Start Menu will remain enabled." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Remove Search Bar
    if ($Opt_RemoveSearchBar) {
        Write-Host "Removing the Search Bar from the Taskbar..." -ForegroundColor Green
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        Set-ItemProperty -Path $regPath -Name "SearchboxTaskbarMode" -Value 0
        Write-Host "Search Bar has been removed from the Taskbar." -ForegroundColor Green
    } else {
        Write-Host "Search Bar will remain on the Taskbar." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Remove Chat icon
    if ($Opt_RemoveChatIcon) {
        Write-Host "Removing the Chat icon from the Taskbar..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "TaskbarMn" -Value 0
        Write-Host "Chat icon has been removed from the Taskbar." -ForegroundColor Green
    } else {
        Write-Host "Chat icon will remain on the Taskbar." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Disable Shake to Minimize
    if ($Opt_DisableShakeMinimize) {
        Write-Host "Disabling the 'Shake to Minimize' feature..." -ForegroundColor Green
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $regPath -Name "DisallowShaking" -Value 1
        Write-Host "'Shake to Minimize' feature has been disabled." -ForegroundColor Green
    } else {
        Write-Host "'Shake to Minimize' feature will remain enabled." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Disable greeting on first login
    if ($Opt_DisableGreeting) {
        Write-Host "Disabling the greeting on first login..." -ForegroundColor Green
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        Set-ItemProperty -Path $regPath -Name "EnableFirstLogonAnimation" -Value 0
        Write-Host "Greeting on first login has been disabled." -ForegroundColor Green
    } else {
        Write-Host "Greeting on first login will remain enabled." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Default to "Do this for all current items"
    if ($Opt_DefaultForAll) {
        Write-Host "Setting default to 'Do this for all current items' in file operations..." -ForegroundColor Green
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\OperationStatusManager"
        New-Item -Path $regPath -Force | Out-Null
        Set-ItemProperty -Path $regPath -Name "ConfirmationCheckBoxDoForAll" -Value 1
        Write-Host "Default for file operations set to 'Do this for all current items'." -ForegroundColor Green
    } else {
        Write-Host "Default for file operations will remain unchanged." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500

    # Add Copy Path to context menu
    if ($Opt_CopyPath) {
        if (!(Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
        }
        if (!(Test-Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath")) {
            New-Item -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Force | Out-Null
        }
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "CanonicalName" -Value "{707C7BC6-685A-4A4D-A275-3966A5A3EFAA}" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "CommandStateHandler" -Value "{3B1599F9-E00A-4BBF-AD3E-B3F99FA87779}" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "Description" -Value "@shell32.dll,-30336" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "Icon" -Value "imageres.dll,-5302" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "InvokeCommandOnSelection" -Value 1 -Type DWord
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "MUIVerb" -Value "@shell32.dll,-30329" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "VerbHandler" -Value "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "VerbName" -Value "copyaspath" -Type String
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shell\windows.copyaspath" -Name "CommandStateSync" -Value "" -Type String

        if (!(Test-Path "HKCR:\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu")) {
            New-Item -Path "HKCR:\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu" -Force | Out-Null
        }
        Set-ItemProperty -Path "HKCR:\AllFilesystemObjects\shellex\ContextMenuHandlers\CopyAsPathMenu" -Name "(Default)" -Value "{f3d06e7c-1e45-4a26-847e-f9fcdee59be0}" -Type String
        Write-Host "'Copy Path' has been added to the context menu." -ForegroundColor Green
    } else {
        Write-Host "'Copy Path' will not be added to the context menu." -ForegroundColor Yellow
    }
    Start-Sleep -Milliseconds 500
    # Restart Explorer to apply changes
    Write-Host "Restarting Windows Explorer to apply changes..." -ForegroundColor Green
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2 # Wait for a moment before restarting Explorer

} else {
    Write-Host "No Windows optimizations will be applied." -ForegroundColor Yellow
}

# Remove default Windows applications
if ($RemoveDefaultAppx) {
    Write-Host "You have chosen to remove default Windows applications." -ForegroundColor Green
    Write-Host "-----------------------------------------"
    Write-Host "Removing Default Windows Applications"
    if ($selectedAppxObjects.Count -gt 0) {
        Write-Host "Removing selected default Windows applications..." -ForegroundColor Green
        foreach ($app in $selectedAppxObjects) {
            Write-Host "Removing $($app.Name)..." -ForegroundColor Cyan
            Get-AppxPackage -PackageTypeFilter Main -Name $app.PackageId | Remove-AppxPackage -ErrorAction SilentlyContinue
            Write-Host "$($app.Name) has been removed." -ForegroundColor Green
        }
    } else {
    Write-Host "No applications selected for removal." -ForegroundColor Yellow
    }
}
else {
    Write-Host "You have chosen not to remove default Windows applications." -ForegroundColor Yellow
}

# Remove OneDrive
if ($OneDriveRemoval) {
    Write-Host "Removing OneDrive..." -ForegroundColor Green
    taskkill /f /im OneDrive.exe | Out-Null
    $uninstallCommand = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe" -Name "UninstallString" -ErrorAction SilentlyContinue).UninstallString
    if ($uninstallCommand) {
        Write-Host "Executing UninstallString: $uninstallCommand"

        # Split the command into executable path and arguments
        # Handles cases where path is quoted or not
        if ($uninstallCommand -match '^"(.+?)"\s*(.*)$') {
            $executable = $Matches[1]
            $arguments  = $Matches[2]
        } elseif ($uninstallCommand -match '^([^\s]+)\s*(.*)$') {
            $executable = $Matches[1]
            $arguments  = $Matches[2]
        } else {
            $executable = $uninstallCommand
            $arguments  = ""
        }

        # Execute the command
        try {
            & "$executable" $arguments
            Write-Host "OneDrive uninstallation command initiated successfully."
        } catch {
            Write-Error "Failed to execute OneDrive uninstall command: $($_.Exception.Message)"
            Write-Host "Please check if the path within the UninstallString is correct and accessible."
        }
    } else {
        Write-Host "UninstallString for OneDrive not found at HKEY_CURRENT_USER path. OneDrive might be uninstalled or installed system-wide."
        Write-Host "Consider checking HKLM paths or uninstalling via Settings if this script fails."
    }
} else {
    Write-Host "You have chosen not to remove OneDrive." -ForegroundColor Yellow
}

# Install selected applications using Winget
if ($InstallSoftware) {
    if ($selectedAppObjects.Count -eq 0) {
        Write-Warning "No valid applications selected. Exiting."
    } else {
        Write-Host "-----------------------------------------"
        Write-Host "You selected the following applications:"
        $selectedAppObjects | ForEach-Object {
            Write-Host "- $($_.DisplayName) ($($_.PackageId))" # Show both for confirmation
        }
        Write-Host "-----------------------------------------"

        # Confirm with the user before proceeding
        $confirm = "Y"
        if ($confirm -eq "Y" -or $confirm -eq "y") {
            Write-Host "Starting installation..."
            foreach ($app in $selectedAppObjects) { # Iterate through the selected objects
                $appId = $app.PackageId # Get the PackageId from the object
                $displayName = $app.DisplayName # Get the DisplayName from the object
                Write-Host "Installing: $displayName ($appId)"
                winget install "$appId" --accept-source-agreements --accept-package-agreements --silent
                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Failed to install $displayName. Winget exit code: $LASTEXITCODE"
                } else {
                    Write-Host "Successfully installed $displayName."
                }
            }
            Write-Host "-----------------------------------------"
            Write-Host "Installation process complete."
        } else {
            Write-Host "Installation cancelled by user."
        }
    }
} else {
    Write-Host "You have chosen not to install any software." -ForegroundColor Yellow
}

# Activate Windows and Office if selected
if($activateProducts) {
    if ($activateWindows) {
    Write-Host "Activating Windows..." -ForegroundColor Green
    & ([ScriptBlock]::Create((irm https://get.activated.win))) /HWID
}
if ($activateOffice) {
    Write-Host "Activating Office..." -ForegroundColor Green
    & ([ScriptBlock]::Create((irm https://get.activated.win))) /Ohook
}
} else {
    Write-Host "You have chosen not to activate any products." -ForegroundColor Yellow
}

# Final message
Write-Host "`n-----------------------------------------" -ForegroundColor Cyan
Write-Host "All selected operations have been completed." -ForegroundColor Cyan
Write-Host "Please restart your computer to apply all changes." -ForegroundColor Cyan
Write-Host "-------------------------------------------" -ForegroundColor Cyan
Read-Host "Press enter to close"
