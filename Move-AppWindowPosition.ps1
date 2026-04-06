[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)][ValidateNotNullOrEmpty()][SupportsWildcards()]
    [Alias('WindowTitle', 'Window', 'Title')]
    [String]$WindowTitleLike = "*snipping*",

    [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [Alias('Screen', 'Destination', 'Monitor', 'MonitorDestination')]
    [Int32]$ScreenDestination = 2,
	
    [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
	[ValidateSet("Full","Left","Right","LeftTop","LeftBottom","RightTop","RightBottom")]
    [Alias('Position', 'Size')]
    [String]$WindowSize = "RightTop",

    [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [Alias('Min', 'Minimized')]
    [Switch]$Minimize
)


#------------------------------------------------------------------------------------------------------------------------
# FUNCTIONS

Function Get-MonitorData ([switch]$ReturnFormsClassValues) {
    # Define values
    [Array] $combined = @()
    [String]$code     = @"
        using System;
        using System.Runtime.InteropServices;
        using System.Collections.Generic;
        using System.Text;

        public class MonitorCollector {
            [StructLayout(LayoutKind.Sequential)]
            public struct RECT {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            public struct MONITORINFOEX {
                public int cbSize;
                public RECT rcMonitor;
                public RECT rcWork;
                public uint dwFlags;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
                public string szDevice;
            }

            public delegate bool MonitorEnumDelegate(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

            [DllImport("user32.dll")]
            public static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumDelegate lpfnEnum, IntPtr dwData);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

            public static List<MONITORINFOEX> GetMonitors() {
                List<MONITORINFOEX> monitors = new List<MONITORINFOEX>();
                EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData) => {
                    MONITORINFOEX mi = new MONITORINFOEX();
                    mi.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
                    if (GetMonitorInfo(hMonitor, ref mi)) {monitors.Add(mi);}
                    return true;
                }, IntPtr.Zero);
                return monitors;
            }
        }
"@

    try {
        # Return Forms class values
        if ($ReturnFormsClassValues) {return [System.Windows.Forms.Screen]::AllScreens | Sort {$_.Bounds.X}}

        # Load type
        try {$type = [MonitorCollector]} catch {$typeMonitorData = Add-Type $code -PassThru}

        # Load monitor data
            # Get data from MonitorCollector class (C# code and WinAPI)
            $apiMonitors = [MonitorCollector]::GetMonitors()

            # Get data from wmi class
            $monitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction SilentlyContinue
            $wmiMonitors = @()
            foreach ($m in $monitors) {
                $wmiMonitors += [PSCustomObject]@{
                    Manufacturer     = ($m.ManufacturerName | ForEach-Object {[char]$_}) -join ''
                    Model            = ($m.UserFriendlyName | ForEach-Object {[char]$_}) -join ''
                    SerialNumber     = ($m.SerialNumberID   | ForEach-Object {[char]$_}) -join ''
                    ProductCodeID    = ($m.ProductCodeID    | ForEach-Object {[char]$_}) -join ''
                    Active           = $m.Active
                    ManufactureDate  = "Year: $($m.YearOfManufacture) Week: $($m.WeekOfManufacture)"
                    InstanceName     = $m.InstanceName
                }
            }

        # Combine found data ($i = id of found monitor)
        $maxCount = [Math]::Max($apiMonitors.Count, $wmiMonitors.Count)
        for ($i = 0; $i -lt $maxCount; $i++) {
            $api = if ($i -lt $apiMonitors.Count) {$apiMonitors[$i]} else {$null}
            $wmi = if ($i -lt $wmiMonitors.Count) {$wmiMonitors[$i]} else {$null}
            $combined += [PSCustomObject]@{
                Device            = if ($api) {$api.szDevice} else {$null}
                X                 = if ($api) {$api.rcMonitor.Left} else {$null}
                Y                 = if ($api) {$api.rcMonitor.Top} else {$null}
                Width             = if ($api) {$api.rcMonitor.Right - $api.rcMonitor.Left} else {$null}
                Height            = if ($api) {$api.rcMonitor.Bottom - $api.rcMonitor.Top} else {$null}
                WorkingAreaX      = if ($api) {$api.rcWork.Left} else {$null}
                WorkingAreaY      = if ($api) {$api.rcWork.Top} else {$null}
                WorkingAreaWidth  = if ($api) {$api.rcWork.Right - $api.rcWork.Left} else {$null}
                WorkingAreaHeight = if ($api) {$api.rcWork.Bottom - $api.rcWork.Top} else {$null}
                IsPrimary         = if ($api) {($api.dwFlags -band 1) -eq 1} else {$false}
                Manufacturer      = if ($wmi) {$wmi.Manufacturer} else {$null}
                Model             = if ($wmi) {$wmi.Model} else {$null}
                SerialNumber      = if ($wmi) {$wmi.SerialNumber} else {$null}
                ProductCodeID     = if ($wmi) {$wmi.ProductCodeID} else {$null}
                ManufactureDate   = if ($wmi) {$wmi.ManufactureDate} else {$null}
                Active            = if ($wmi) {$wmi.Active} else {$false}
                WmiInstance       = if ($wmi) {$wmi.InstanceName} else {$null}
            }
        }

        # Return value
        return $combined
    } catch {Write-Host ($_ | Out-String).Trim()}
}

Function Move-WindowToScreen {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)][ValidateNotNullOrEmpty()][SupportsWildcards()]
        [Alias('WindowTitle', 'Window', 'Title')]
        [String]$WindowTitleLike,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('TimeOut')]
        [Int32]$TimeOutInSeconds = 5,

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('Screen', 'Destination', 'Monitor', 'MonitorDestination')]
        [Int32]$ScreenDestination = 1,
	
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
	    [ValidateSet("Full","Left","Right","LeftTop","LeftBottom","RightTop","RightBottom")]
        [Alias('Position', 'Size')]
        [String]$WindowSize = "Full",
	
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
	    [ValidateSet("All","Results","None")]
        [Alias('Console', 'Output')]
        [String]$ConsoleOutput = "All",

        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [Alias('Min', 'Minimized')]
        [Switch]$Minimize,
	
	    [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
	    [Switch]$ScreenInfoFormsClass
    )
	
	try {
        # Define values
        [String]$codeProc = @"
            using System;
            using System.Diagnostics;
            using System.Collections.Generic;
            using System.Text.RegularExpressions;

            public class ProcessFinder {
                public static List<Process> GetByTitleWildcard(string wildcard) {
                    List<Process> found = new List<Process>();
                    string pattern = "^" + Regex.Escape(wildcard).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
                    Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
                    foreach (Process p in Process.GetProcesses()) {
                        try {
                            if (p.MainWindowHandle != IntPtr.Zero) {
                                if (regex.IsMatch(p.MainWindowTitle)) {
                                    found.Add(p);
                                }
                            }
                        } catch {
                            continue;
                        }
                    }
                    return found;
                }
            }
"@

        [String]$codeWinApi = @"
            using System;
            using System.Text;
            using System.Runtime.InteropServices;

            namespace WinAPI {
                public enum ShowWindowCommands {
                    SW_HIDE = 0,
                    SW_SHOWNORMAL = 1,
                    SW_SHOWMINIMIZED = 2,
                    SW_SHOWMAXIMIZED = 3,
                    SW_SHOWNOACTIVATE = 4,
                    SW_SHOW = 5,
                    SW_MINIMIZE = 6,
                    SW_SHOWMINNOACTIVE = 7,
                    SW_SHOWNA = 8,
                    SW_RESTORE = 9,
                    SW_SHOWDEFAULT = 10,
                    SW_FORCEMINIMIZE = 11
                }

                [StructLayout(LayoutKind.Sequential)]
                public struct RECT {
                    public int Left;
                    public int Top;
                    public int Right;
                    public int Bottom;
                }

                public class Win32 {
                    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
                        
                    [DllImport("user32.dll")]
                    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

                    [DllImport("user32.dll")]
                    public static extern bool ShowWindow(IntPtr hWnd, ShowWindowCommands nCmdShow);

                    [DllImport("user32.dll")]
                    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

                    [DllImport("user32.dll")]
                    public static extern bool SetForegroundWindow(IntPtr hWnd);

                    [DllImport("user32.dll")]
                    [return: MarshalAs(UnmanagedType.Bool)]
                    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

                    [DllImport("user32.dll")]
                    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

                    [DllImport("user32.dll")]
                    public static extern int GetWindowTextLength(IntPtr hWnd);

                    [DllImport("user32.dll")]
                    public static extern bool IsWindowVisible(IntPtr hWnd);
                }
            }
"@

        # Load types
        try {if (Get-Variable -Name 'typeProc' -Scope 'Global') {if ($ConsoleOutput -in @("All")) {Write-Host "ProcessFinder already existed."}}}
        catch {$global:typeProc = Add-Type -TypeDefinition $codeProc -PassThru; if ($ConsoleOutput -in @("All")) {Write-Host "ProcessFinder loaded."}}

        try {if (Get-Variable -Name 'typeWinApi' -Scope 'Global') {if ($ConsoleOutput -in @("All")) {Write-Host "WinAPI already existed."}}}
        catch {$global:typeWinApi = Add-Type -TypeDefinition $codeWinApi -PassThru; if ($ConsoleOutput -in @("All")) {Write-Host "WinAPI loaded."}}

        # Load window handles and search for window title
        [Int32] $global:foundHandle = 0
        [Int32] $i                  = 0
        [String]$foundWith          = [String]::Empty
        while ($true) {
            if ($i -ge ($TimeOutInSeconds*2)) {break}

            # ProcessFinder (C#-API)
            $procs = [ProcessFinder]::GetByTitleWildcard($WindowTitleLike)
            if ($procs) {if ($procs.Count -gt 0) {
                $windowHandles = ($procs | Select-Object -Property "MainWindowHandle").MainWindowHandle
                $foundWith = "ProcessFinder (C#-API)"
                break
            }}

            # Win32-API
            $callback = [WinAPI.Win32+EnumWindowsProc]{
                param($hWnd, $lParam)

                if (-not [WinAPI.Win32]::IsWindowVisible($hWnd)) {return $true}

                $length = [WinAPI.Win32]::GetWindowTextLength($hWnd)
                if ($length -eq 0) {return $true}
                $builder = New-Object System.Text.StringBuilder $length+1
                [WinAPI.Win32]::GetWindowText($hWnd, $builder, $builder.Capacity)
                $title = $builder.ToString()

                if ($title -like $WindowTitleLike) {
                    $global:foundHandle = $hWnd
                    return $false  # Stop enumeration
                }

                return $true
            }
            ([WinAPI.Win32])::EnumWindows($callback, [IntPtr]::Zero) | Out-Null
            if ($global:foundHandle -ne 0) {
                $windowHandles = @($global:foundHandle)
                $foundWith = "Win32-API"
                break
            }
            
            $i++
            Start-Sleep -Milliseconds 500
        }

        # Evaluate searching result
        if ($i -ge ($TimeOutInSeconds*2)) {return}
        if ($windowHandles) {if ($ConsoleOutput -in @("All","Results")) {
            Write-Host (
                "Found window handles  : $($windowHandles)`n" + 
                "    Window Title like : $($WindowTitleLike)`n" + 
                "    Method            : $($foundWith)"
            )
        }}

        # Stop if no window handles are found
        if (!($windowHandles)) {if ($ConsoleOutput -in @("All","Results")) {Write-Host "No windows handles found for '$WindowTitleLike'."}; return}

        # Load monitor statistics
        if ($ScreenInfoFormsClass) {
            $screens = Get-MonitorData -ReturnFormsClassValues
            if ([int]$ScreenDestination -lt 1) {[int]$ScreenDestination = 1}
            if ([int]$ScreenDestination -gt $screens.Count) {[int]$ScreenDestination = $screens.Count}
            $screenSelected = $screens[$ScreenDestination-1]
            $screenPosX = $screenSelected.WorkingArea.X
            $screenPosY = $screenSelected.WorkingArea.Y
            $screenSizeWidth = $screenSelected.WorkingArea.Width
            $screenSizeHeight = $screenSelected.WorkingArea.Height

            if ($ConsoleOutput -in @("All")) {
                Write-Host "Screens found         : $($screens.Count)"
                foreach ($screen in $screens) {
                    Write-Host (
                        "    DeviceName        : $($screen.DeviceName)`n" + 
                        "        Full Screen`n" + 
                        "            Positon X : $($screen.Bounds.X)`n" + 
                        "            Positon Y : $($screen.Bounds.Y)`n" + 
                        "            Width     : $($screen.Bounds.Width)`n" + 
                        "            Height    : $($screen.Bounds.Height)`n" + 
                        "        WorkingArea`n" + 
                        "            Positon X : $($screen.WorkingArea.X)`n" + 
                        "            Positon Y : $($screen.WorkingArea.Y)`n" + 
                        "            Width     : $($screen.WorkingArea.Width)`n" + 
                        "            Height    : $($screen.WorkingArea.Height)`n" + 
                        "        Primary       : $($screen.Primary)`n" + 
                        "        BitsPerPixel  : $($screen.BitsPerPixel)"
                    )
                }
            }
        }
        else {
            $screens = Get-MonitorData | ? {$_.Device}
            if ([int]$ScreenDestination -lt 1) {[int]$ScreenDestination = 1}
            if ([int]$ScreenDestination -gt $screens.Count) {[int]$ScreenDestination = $screens.Count}
            $screenSelected = $screens[$ScreenDestination-1]
            $screenPosX = $screenSelected.WorkingAreaX
            $screenPosY = $screenSelected.WorkingAreaY
            $screenSizeWidth = $screenSelected.WorkingAreaWidth
            $screenSizeHeight = $screenSelected.WorkingAreaHeight

            if ($ConsoleOutput -in @("All")) {
                Write-Host "Screens found         : $($screens.Count)"
                foreach ($screen in $screens) {
                    Write-Host (
                        "    Device            : $($screen.Device)`n" + 
                        "        Full Screen`n" + 
                        "            Positon X : $($screen.X)`n" + 
                        "            Positon Y : $($screen.Y)`n" + 
                        "            Width     : $($screen.Width)`n" + 
                        "            Height    : $($screen.Height)`n" + 
                        "        WorkingArea`n" + 
                        "            Positon X : $($screen.WorkingAreaX)`n" + 
                        "            Positon Y : $($screen.WorkingAreaY)`n" + 
                        "            Width     : $($screen.WorkingAreaWidth)`n" + 
                        "            Height    : $($screen.WorkingAreaHeight)`n" + 
                        "        Primary       : $($screen.IsPrimary)`n" + 
                        "        Manufacturer  : $($screen.Manufacturer)"
                    )
                }
            }
        }
        $repaint = $true

        if ($ConsoleOutput -in @("All")) {
            Write-Host (
                "Selected screen       : $($ScreenDestination)`n" + 
                "    Positon X         : $($screenPosX)`n" + 
                "    Positon Y         : $($screenPosY)`n" + 
                "    Width             : $($screenSizeWidth)`n" + 
                "    Height            : $($screenSizeHeight)`n" + 
                "Selected WindowSize   : $($WindowSize)"
            )
        }

        # Modify Window
        foreach ($handle in $windowHandles) {
            if (!($handle -and ($handle -ne 0))) {Continue}

            $wndRect       = [WinAPI.RECT]::new()
            $wndRectLoaded = [WinAPI.Win32]::GetWindowRect($handle, [ref]$wndRect)
            if ($wndRectLoaded -ne $true) {Continue}

            if ($ConsoleOutput -in @("All")) {
                Write-Host (
                    "Window handle         : $($handle)`n" + 
                    "    Active settings`n" + 
                    "        Left          : $($wndRect.Left)`n" + 
                    "        Top           : $($wndRect.Top)`n" + 
                    "        Width         : $($wndRect.Right  - $wndRect.Left)`n" + 
                    "        Height        : $($wndRect.Bottom - $wndRect.Top)"
                )
            }

			Switch ($WindowSize) {
				"Full" {
					$wndLeft     = $screenPosX
					$wndTop      = $screenPosY
					$wndWidth    = $wndRect.Right  - $wndRect.Left
					$wndHeight   = $wndRect.Bottom - $wndRect.Top
				}
				"Left" {
					$wndLeft     = $screenPosX
					$wndTop      = $screenPosY
					$wndWidth    = $screenSizeWidth/2
					$wndHeight   = $screenSizeHeight
				}
				"Right" {
					$wndLeft     = $screenPosX + ($screenSizeWidth/2)
					$wndTop      = $screenPosY
					$wndWidth    = $screenSizeWidth/2
					$wndHeight   = $screenSizeHeight
				}
				"LeftTop" {
					$wndLeft     = $screenPosX
					$wndTop      = $screenPosY
					$wndWidth    = $screenSizeWidth/2
					$wndHeight   = $screenSizeHeight/2
				}
				"LeftBottom" {
					$wndLeft     = $screenPosX
					$wndTop      = $screenPosY + ($screenSizeHeight/2)
					$wndWidth    = $screenSizeWidth/2
					$wndHeight   = $screenSizeHeight/2
					
				}
				"RightTop" {
					$wndLeft     = $screenPosX + ($screenSizeWidth/2)
					$wndTop      = $screenPosY
					$wndWidth    = $screenSizeWidth/2
					$wndHeight   = $screenSizeHeight/2
				}
				"RightBottom" {
					$wndLeft     = $screenPosX + ($screenSizeWidth/2)
					$wndTop      = $screenPosY + ($screenSizeHeight/2)
					$wndWidth    = $screenSizeWidth/2
					$wndHeight   = $screenSizeHeight/2
				}
			}

            $windowShowChange          = [WinAPI.Win32]::ShowWindow($handle, [WinAPI.ShowWindowCommands]::SW_HIDE)
            $windowShowChange          = [WinAPI.Win32]::ShowWindow($handle, [WinAPI.ShowWindowCommands]::SW_SHOWNOACTIVATE)
            $windowPosSizeChange       = [WinAPI.Win32]::MoveWindow($handle, $wndLeft, $wndTop, $wndWidth, $wndHeight, $repaint)
            if ($Minimize) {
                $windowShowChange          = [WinAPI.Win32]::ShowWindow($handle, [WinAPI.ShowWindowCommands]::SW_MINIMIZE)}
            else {
                if ($WindowSize -eq "Full") {
                    $windowShowChange          = [WinAPI.Win32]::ShowWindow($handle, [WinAPI.ShowWindowCommands]::SW_SHOWNORMAL)
                    $windowShowChange          = [WinAPI.Win32]::ShowWindow($handle, [WinAPI.ShowWindowCommands]::SW_SHOWMAXIMIZED)
                    $windowToForegroundSuccess = [WinAPI.Win32]::SetForegroundWindow($handle)
                }
                else {
                    $windowShowChange          = [WinAPI.Win32]::ShowWindow($handle, [WinAPI.ShowWindowCommands]::SW_SHOWNORMAL)
                    $windowToForegroundSuccess = [WinAPI.Win32]::SetForegroundWindow($handle)
                }
            }

            if ($ConsoleOutput -in @("All","Results")) {
                Write-Host (
                    "    New settings`n" + 
                    "        Left          : $($wndLeft)`n" + 
                    "        Top           : $($wndTop)`n" + 
                    "        Width         : $($wndWidth)`n" + 
                    "        Height        : $($wndHeight)`n" + 
                    "        Minimize      : $($Minimize)"
                )
            }
            
        }
    } catch {Write-Host ($_ | Out-String).Trim()}
}


#------------------------------------------------------------------------------------------------------------------------
# BEGIN

    # Preparation
    $ErrorActionPreference = 'Stop'
    Clear-Host

    # Start process
    Move-WindowToScreen -WindowTitleLike $WindowTitleLike -ScreenDestination $ScreenDestination -WindowSize $WindowSize