# ===================================================
# On a new machine you will need to run:
# Set-ExecutionPolicy RemoteSigned
#
# Instructions:
# 1. Download this file
#       Invoke-WebRequest -Uri https://raw.githubusercontent.com/BrassStack/NewMachine/master/FirstInstall.ps1 -Method Get -OutFile FirstInstall.ps1
# 2. Execute it
#       ./FirstInstall.ps1
# ===================================================

[CmdletBinding()]
param (
    [switch] $Force,
    [switch] $Chocolatey
)

# ===================================================
# Run as administrator if not already
# ===================================================
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
        Exit
    }
}

# Checks $Path then runs a command to populate that path, optionally clearing it out first
function RestoreItem() {
    [CmdletBinding()]
    param (
        [string] $Path,
        [string] $AddCmd
    )

    if ( Test-Path $Path ) {
        if ( $Force ) {
            Write-Host -ForegroundColor magenta "`nReplacing $Path`n"
            Remove-Item -Force -Recurse $Path
            Invoke-Expression $AddCmd
        }
        else {
            Write-Host -ForegroundColor white "Skipped existing $Path"
        }
    }
    else {
        Write-Host -ForegroundColor cyan "`nCloning to $Path`n"
        Invoke-Expression $AddCmd
    }
    if ( -not $? ) {
        throw "Cloning to $Path failed"
    }
}

# Waits for user to press a key to continue, but allows for ^C to cancel
function Wait-KeyTimeout() {
    [CmdletBinding()]
    param(
        [string] $Message = "Press any key to continue...",
        [int] $Seconds = 60,
        [ConsoleColor] $Color = [ConsoleColor]::Green,
        [switch] $NoTimeout,
        [switch] $KeepMessage
    )

    Start-Sleep -Milliseconds 200
    $host.ui.RawUI.FlushInputBuffer()

    $count = if ( $NoTimeout ) { "" } else { $Seconds }
    $i = 0
    while ( $NoTimeout -or $count -gt 0) { 
        Write-Host -ForegroundColor $Color -NoNewline ("`r{0} $Message $count  " -f '/-\|'[($i++ % 4)])
        Start-Sleep -Milliseconds 250 
        if ( -not $NoTimeout -and $i % 4 -eq 0 ) { $count-- }
        if ( $host.UI.RawUI.KeyAvailable ) {
            $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ( $key.VirtualKeyCode -ne 17 ) {
                break
            }
        }
    }
    if ( -not $KeepMessage ) {
        Write-Host -NoNewline "`r" (" " * ($Message.Length + 5))
    }
}

Write-Host " "
Write-Host -ForegroundColor Green "**** NEW MACHINE SETUP! ****"
Write-Host " "
Write-Host -ForegroundColor Yellow "Warning: This script does a lot of installs and tweaks. Some of these settings could be dangerous in uncontrolled environments. Be sure you understand the significance of the following:"
Write-Host " "
Write-Host "1. All current public networks will be set to private"
Write-Host "2. PSGallery will be set as Trusted"
Write-Host "3. NuGet will be added as a package provider"
Write-Host "4. PS Remoting will be enabled with CredSSP for all machines"
Write-Host "5. Both PowerShell directories will be replaced"
Write-Host "6. Internet Zone Security will be modified"
Write-Host " "
Write-Host -ForegroundColor Yellow "If you do not fully understand any one of these items, Press Control+C NOW."
Write-Host " "
Wait-KeyTimeout

try {
    
    Write-Host -ForegroundColor cyan "`nBeginning new machine setup ...`n"
    $Error.Clear()
    
    # Make sure default repo is trusted and install useful modules
    Write-Host -ForegroundColor cyan "`nSetting up Windows PowerShell ...`n"
    Get-NetConnectionProfile -NetworkCategory Public -ErrorAction SilentlyContinue | Set-NetConnectionProfile -NetworkCategory Private -ErrorAction SilentlyContinue
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers -ForceBootstrap
    Install-Module posh-git -Scope AllUsers -Force
    Install-Module PowerShellGet -Scope AllUsers -Force
    Get-WindowsCapability -Online -Name Rsat.ActiveDirectory* | Add-WindowsCapability -Online -ErrorAction SilentlyContinue

    Enable-PSRemoting
    Enable-WSManCredSSP -Role Client -DelegateComputer * -Force
    
    # ===================================================
    # Global items
    # ===================================================

    # Install Chocolatey and git
    if ( $Chocolatey ) {
        Get-Command choco -ErrorAction SilentlyContinue
        if ( -not $? ) {
            Write-Host -ForegroundColor cyan "`nInstalling Chocolatey...`n"
            Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
            if ( -not $? ) {
                throw "Installing Chocolatey failed"
            }
        }
        else {
            Write-Host -ForegroundColor cyan "`nChocolatey is already installed...`n"
        }
        choco feature enable -n=allowGlobalConfirmation
    }

    Get-Command git -ErrorAction SilentlyContinue
    if ( -not $? ) {
        Write-Host -ForegroundColor cyan "`nInstalling Git for Windows...`n"
        if ( $Chocolatey ) {
            choco install git --limitoutput
        }
        else {
            winget install Git.Git --disable-interactivity --accept-package-agreements --accept-source-agreements
        }
        if ( -not $? ) {
            throw "Installing Git for Windows failed"
        }
    }
    else {
        Write-Host -ForegroundColor cyan "`Git for Windows is already installed...`n"
    }

    # Refresh environment so git can be used
    if ( $Chocolatey ) {
        $env:ChocolateyInstall = Convert-Path "$((Get-Command choco).Path)\..\.."   
        Import-Module "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
        refreshenv
    }
    else {
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }

    # Clone PowerShell v5 profile if profile workspace is not in it yet
    $ps5path = "$HOME\Documents\WindowsPowerShell"
    if ( -not (Test-Path $ps5path\My-Profile.code-workspace) ) {
        Write-Host -ForegroundColor cyan "`nCloning to $ps5path`n"
        Set-Location $HOME
        Remove-Item -Force -Recurse $ps5path -ErrorAction SilentlyContinue
        git clone https://github.com/BrassStack/PowerShell.git $ps5path
    }
    else {
        Write-Host -ForegroundColor white "`nSkipped existing $ps5path"
    }
    if ( -not $? ) {
        throw "Cloning ps5 profile failed"
    }


    # ===================================================
    # Full set of backup settings and files now available
    # ===================================================
    
    Set-Location $ps5path
    
    Write-Host -ForegroundColor cyan "`nApplying Windows Tweaks...`n"

    # Restore mouse cursors
    Write-Host -ForegroundColor white "Restoring mouse cursors..."
    Copy-Item ./Backup/Cursors/* C:\windows\Cursors\ -Recurse -ErrorAction Continue
    if ( $? ) {
        reg import ./Cursors.reg
    } else {
        Write-Host -ForegroundColor darkyellow "Cursor copy failed"
    }

    # Find a way to lock monitor scaling at 100%

    # Set Dvorak
    Write-Host -ForegroundColor white "Installing DVORAK..."
    $en = (Get-WinUserLanguageList)[0]
    $en.InputMethodTips.Insert( 0, "0409:00010409" )
    Set-WinUserLanguageList $en -Force -ErrorAction Continue
    if ( -not $? ) {
        Write-Host -ForegroundColor darkyellow "DVORAK installation failed, adding registry keys instead ..."
        reg import ./Keyboard.reg
    }

    # Reconfigure IE security
    Write-Host -ForegroundColor white "Configuring Internet security..."
    reg import ./TrustedSitesZoneSecurity.reg

    # Add power profile (requires full path)
    Write-Host -ForegroundColor white "Adding power profile ..."
    $powerGuid = New-Guid
    powercfg -import "$HOME/Documents/WindowsPowerShell/Backup/balanced.pow" $powerGuid
    if ( -not $? ) {
        powercfg -setactive $powerGuid
    } else {
        # If import failed, settle for resetting the timeouts
        powercfg /X monitor-timeout-ac 0
        powercfg /X standby-timeout-ac 0
        powercfg /X monitor-timeout-dc 10
        powercfg /X standby-timeout-dc 30
    }

    Write-Host -ForegroundColor white "Applying UI tweaks ..."

    # Set keyboard repeat delay to zero
    Set-ItemProperty "HKCU:\Control Panel\Keyboard\" -Name KeyboardDelay -Value 0

    # Set recycle bin drive space usage
    $volume = mountvol C:\ /L 
    $guid = [regex]::Matches( $volume, '{([-0-9A-F].*?)}' )
    $regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\BitBucket\Volume\"+$guid.value
    Set-ItemProperty -Path $regKey -Name "MaxCapacity" -Type DWord -Value 1024
    
    # Turn on Developer mode
    reg add "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" /t REG_DWORD /f /v "AllowDevelopmentWithoutDevLicense" /d "1"

    # Tweaks from https://github.com/Disassembler0/Win10-Initial-Setup-Script
    Import-Module ./Tweaks.psm1
    DisableAeroShake
    DisableAccessibilityKeys
    ShowTaskManagerDetails
    ShowFileOperationsDetails
    HideTaskbarSearch
    SetTaskbarCombineWhenFull
    DisableShortcutInName
    DisableThumbnailCache
    DisableThumbsDBOnNetwork
    SetAppsDarkMode
    SetSystemDarkMode
    ShowExplorerTitleFullPath
    ShowKnownExtensions
    ShowNavPaneLibraries
    EnableRestoreFldrWindows
    ShowSelectCheckboxes
    HideNetworkFromExplorer
    Hide3DObjectsFromThisPC
    UninstallMsftBloat
    DisableXboxFeatures
    DisableEdgePreload
    DisableIEFirstRun
    UninstallXPSPrinter
    RemoveFaxPrinter
    UnpinStartMenuTiles
    UnpinTaskbarIcons
    
    # Repository paths
    $ps7path = "$HOME\Documents\PowerShell"
    $gitpath = "$HOME\.config\git"
    $ahkpath = "$HOME\AutoHotkey\Scripts"
    
    # Add PowerShell v7 worktree
    RestoreItem $ps7path "git worktree add $ps7path v7"

    # Clone gitconfig
    RestoreItem $gitpath "git clone https://github.com/BrassStack/gitconfig.git $gitpath"

    # Clone autohotkey
    RestoreItem $ahkpath "git clone https://github.com/BrassStack/AutoHotkey.git $ahkpath"

    # Restore Equalizer APO config
    Write-Host -ForegroundColor cyan "`nRestoring EqualizerAPO config...`n"
    New-Item -ItemType Directory "C:\Program Files\EqualizerAPO" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "C:\Program Files\EqualizerAPO\config" -ErrorAction SilentlyContinue
    ./Backup/EqualizerAPO64-1.3.exe
    Copy-Item ./Backup/EQAPOconfig.txt "C:\Program Files\EqualizerAPO\config\config.txt" -Verbose -Force

    # Create some directories
    if ( (Test-Path "$env:OneDrive\Documents") -and -not (Test-Path "$HOME\Documents") ) {
        Write-Host -ForegroundColor cyan "`nAdding short paths to libraries...`n"
        New-Item -ItemType Junction -Path "$HOME\Documents" -Target "$env:OneDrive\Documents" -ErrorAction SilentlyContinue
        New-Item -ItemType Junction -Path "$HOME\Pictures" -Target "$env:OneDrive\Pictures" -ErrorAction SilentlyContinue
        New-Item -ItemType Junction -Path "$HOME\Desktop" -Target "$env:OneDrive\Desktop" -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory "$HOME\Source" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "$HOME\Source\Demos" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "$HOME\Source\Github" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "$HOME\Source\Libraries" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "$HOME\Source\PowerShell" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "$HOME\Source\Utilities" -ErrorAction SilentlyContinue
    New-Item -ItemType Directory "$HOME\Source\WebApps" -ErrorAction SilentlyContinue
    
    Write-Host -ForegroundColor cyan "`nAdding Projects library...`n"
    Copy-Item "./Backup/Projects.library-ms" "$HOME\AppData\Roaming\Microsoft\Windows\Libraries\" -Verbose -ErrorAction Continue
    
    # Import scheduled task xmls
    Write-Host -ForegroundColor cyan "`nRestoring scheduled tasks...`n" 
    $un = (whoami)
    Register-ScheduledTask -xml (Get-Content './Backup/Scheduled Tasks/Start AHK.xml' | Out-String) -TaskName "Start AHK" -TaskPath "\" -User $un -Force
    Register-ScheduledTask -xml (Get-Content './Backup/Scheduled Tasks/Restore Windows after reboot.xml' | Out-String) -TaskName "Restore Windows after reboot" -TaskPath "\" -User $un -Force
    Register-ScheduledTask -xml (Get-Content './Backup/Scheduled Tasks/Save Open Windows on Restart event.xml' | Out-String) -TaskName "Save Open Windows on Restart event" -TaskPath "\" -User $un -Force


    # ===================================================
    # Add IWU repo and set up work-specific items
    # ===================================================
    ./Install-WorkItems.ps1


    # ===================================================
    # Final Messages and launch VS Code with profile
    # ===================================================
    Write-Host -ForegroundColor cyan "`nInstalling VS Code...`n"
    if ( $Chocolatey ) {
        choco install vscode --limitoutput
    }
    else {
        winget install Microsoft.VisualStudioCode --scope machine
    }
    if ( -not $? ) {
        throw "Installing VS Code failed"
    }

    Write-Host -ForegroundColor green "`n`nProfile data cloned!"

    Start-Process "$HOME\Documents\WindowsPowerShell\My-Profile.code-workspace"

    Write-Host -ForegroundColor Cyan "Now installing applications."
    Wait-KeyTimeout


    # ===================================================
    # Install applications
    # ===================================================

    if ( Test-Path "\\iwufiles\common\uit\datateam" ) {
        if ( $Chocolatey ) {
            choco install choco-work.config
        }
        else {
            winget import -i .\winget.pkgs.json --disable-interactivity --accept-package-agreements --accept-source-agreements
        }
    }
    else {
        if ( $Chocolatey ) {
            choco install choco-home.config
        }
        else {
            winget import -i .\winget.pkgs.json --disable-interactivity --accept-package-agreements --accept-source-agreements
        }
    }

    Write-Host -ForegroundColor Cyan "Now installing VS Extensions"
    Wait-KeyTimeout
    ./Install-VsExtensions.ps1
}
catch {
    Write-Host -ForegroundColor red "Aborted:"
    $Error | Out-File -FilePath Firstinstall.log
    $_
    # If the script re-ran itself as Administrator, the new window will close before you get to see the output unless we wait
    Wait-KeyTimeout "Press any key to exit." -NoTimeout -Color Red
}
