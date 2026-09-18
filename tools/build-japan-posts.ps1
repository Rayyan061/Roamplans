param(
    [string]$SourceDir = 'D:\Roamplans 11 categories\Japan Travel Guide\Japan Trip Planning',
    [string]$OutputDir = 'D:\Roamplans\japan-trip-planning',
    [string]$Template = 'D:\Roamplans\travel-insurance\family-travel-insurance-guide.html'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Html([string]$s) { if ($null -eq $s) { return '' }; return [Net.WebUtility]::HtmlEncode($s) }
function Slug([string]$s) {
    $v = $s.ToLowerInvariant() -replace '[\u2018\u2019]', "'"
    $v = $v -replace "[^a-z0-9]+", '-'
    return $v.Trim('-')
}
function Description([string]$title, [string]$first) {
    $s = ($first -replace '\s+', ' ').Trim()
    if ($s.Length -gt 150) {
        $cut = $s.Substring(0,150)
        $lastSpace = $cut.LastIndexOf(' ')
        if ($lastSpace -gt 110) { $cut = $cut.Substring(0,$lastSpace) }
        $s = $cut.TrimEnd(' ',',',';',':','.') + '.'
    }
    if ($s.Length -lt 80) { $s = "${title}: practical advice for first-time visitors planning dates, routes, costs, packing, and a smoother Japan trip." }
    return $s
}
function Read-Docx([string]$path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        $reader = [IO.StreamReader]::new($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $items = [Collections.Generic.List[object]]::new()
    foreach ($node in $xml.SelectNodes('//w:body/*',$ns)) {
        if ($node.LocalName -eq 'p') {
            $styleNode = $node.SelectSingleNode('./w:pPr/w:pStyle',$ns)
            $style = if ($styleNode) { $styleNode.GetAttribute('val','http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { '' }
            $isList = $null -ne $node.SelectSingleNode('./w:pPr/w:numPr',$ns)
            $parts = [Collections.Generic.List[string]]::new()
            foreach ($run in $node.SelectNodes('.//w:r',$ns)) {
                $txt = (($run.SelectNodes('.//w:t',$ns) | ForEach-Object { $_.InnerText }) -join '')
                $hasBreak = $null -ne $run.SelectSingleNode('./w:br',$ns)
                if (-not $txt -and -not $hasBreak) { continue }
                $safe = if($txt){ Html $txt }else{ '' }
                if ($run.SelectSingleNode('./w:rPr/w:b',$ns)) { $safe = "<strong>$safe</strong>" }
                if ($run.SelectSingleNode('./w:rPr/w:i',$ns)) { $safe = "<em>$safe</em>" }
                if ($run.SelectSingleNode('./w:br',$ns)) { $safe += '<br>' }
                $parts.Add($safe)
            }
            $html = ($parts -join '')
            $plain = (($node.SelectNodes('.//w:t',$ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
            if ($plain) { $items.Add([pscustomobject]@{Type='p';Style=$style;IsList=$isList;Text=$plain;Html=$html}) }
        } elseif ($node.LocalName -eq 'tbl') {
            $rows = [Collections.Generic.List[object]]::new()
            foreach ($tr in $node.SelectNodes('./w:tr',$ns)) {
                $cells = @()
                foreach ($tc in $tr.SelectNodes('./w:tc',$ns)) {
                    $cells += ((($tc.SelectNodes('.//w:t',$ns) | ForEach-Object { $_.InnerText }) -join ' ').Trim())
                }
                $rows.Add($cells)
            }
            $items.Add([pscustomobject]@{Type='table';Rows=$rows})
        }
    }
    return $items
}

$templateText = Get-Content -LiteralPath $Template -Raw -Encoding UTF8
$styleMatch = [regex]::Match($templateText,'(?s)<style>(.*?)</style>')
if (-not $styleMatch.Success) { throw 'Could not locate the locked layout CSS.' }
$css = $styleMatch.Groups[1].Value

# Small Japan-specific additions; the locked visual system remains unchanged.
$css += @'

.rp-hero .hero-inner:after{content:"JP";position:absolute;right:0;bottom:-38px;font-family:Georgia,"Times New Roman",serif;font-size:150px;font-weight:700;line-height:1;color:rgba(255,255,255,.07);pointer-events:none}.rp-step{margin:28px 0;padding:22px 24px;border-left:5px solid var(--rp-gold);border-radius:0 18px 18px 0;background:var(--rp-sand)}.rp-step strong{color:var(--rp-ink)}
.rp-faq-question{list-style:none}.rp-faq-question::-webkit-details-marker{display:none}.rp-faq-item[open] .rp-faq-icon:after{transform:translate(-50%,-50%) rotate(0deg)}.rp-faq-item[open] .rp-faq-answer{grid-template-rows:1fr}
'@

$posts = @(
    @{N=1; Slug='how-to-plan-a-trip-to-japan'; Short='Plan a Japan trip'; Desc='Plan your first Japan trip step by step, from choosing dates, trip length, budget, and cities to booking flights, hotels, transport, and activities.'; Meta=@('Step-by-step route','First-time friendly','Booking timeline'); Note='Build the route before booking.'; NoteText='Start with usable days, dates, budget, and logical bases. Flights and hotels become easier once those four decisions are clear.'},
    @{N=2; Slug='how-many-days-do-you-need-in-japan'; Short='How many days in Japan?'; Desc='Compare realistic 7-, 10-, and 14-day Japan itineraries, including usable sightseeing days, city bases, transfer time, and pacing for first visits.'; Meta=@('7, 10 & 14 days','Realistic pacing','Route examples'); Note='Count usable days, not nights.'; NoteText='Arrival, departure, and city-transfer days rarely work like full sightseeing days. Protect enough time for your main bases.'},
    @{N=3; Slug='best-time-to-visit-japan'; Short='Best time for Japan'; Desc='Find the best time to visit Japan by comparing seasonal weather, crowds, travel costs, scenery, events, and month-by-month planning trade-offs.'; Meta=@('Weather by season','Crowds & costs','Month-by-month help'); Note='Choose the season for your priorities.'; NoteText='Weather, crowds, prices, and scenery move together. Compare the complete trade-off instead of choosing from photos alone.'},
    @{N=4; Slug='japan-trip-cost'; Short='Japan trip cost'; Desc='Estimate your Japan trip cost for 7, 10, or 14 days with practical budget ranges for flights, hotels, transport, food, attractions, and extras.'; Meta=@('7, 10 & 14 days','Budget ranges','Cost breakdown'); Note='Let the route shape the budget.'; NoteText='Flights, hotel level, intercity travel, and pace create the biggest differences. Keep a separate emergency buffer.'},
    @{N=5; Slug='what-to-pack-for-japan'; Short='Japan packing list'; Desc='Use this first-time Japan packing list to choose seasonal clothing, shoes, travel documents, electronics, toiletries, medication, and luggage.'; Meta=@('First-time checklist','Seasonal layers','Luggage strategy'); Note='Pack for movement and the season.'; NoteText='Japan trips often combine long walking days, trains, stairs, and compact rooms. A lighter, flexible bag is easier to manage.'},
    @{N=6; Slug='japan-travel-mistakes'; Short='Japan travel mistakes'; Desc='Avoid common first-time Japan travel mistakes involving rushed routes, train planning, luggage, cash, etiquette, reservations, and daily pacing.'; Meta=@('First-trip pitfalls','Smarter planning','Easy fixes'); Note='A simpler trip is usually better.'; NoteText='Too many bases, rigid days, and unpriced transport create avoidable stress. Leave room for normal travel changes.'}
)

$category = 'Japan Trip Planning'
$sourceLeaf = Split-Path -Leaf $SourceDir
if ($sourceLeaf -eq 'Japan Itineraries') {
    $category = 'Japan Itineraries'
    $posts = @(
        @{N=1; Slug='japan-7-day-itinerary'; Short='7 days in Japan'; Desc='Follow a realistic 7-day Japan itinerary for first-time visitors, with a practical Tokyo, Kyoto, and Osaka route, daily plans, transport, and pacing.'; Meta=@('7-day route','First-time friendly','Day-by-day plan'); Note='Keep one week focused.'; NoteText='Seven days work best with a small number of bases. Protect sightseeing time by avoiding unnecessary hotel changes and backtracking.'},
        @{N=2; Slug='japan-10-day-itinerary'; Short='10 days in Japan'; Desc='Plan a balanced 10-day Japan itinerary for first-time visitors, including Tokyo, Kyoto, Osaka, route order, transport days, and realistic daily pacing.'; Meta=@('10-day route','Tokyo, Kyoto & Osaka','Balanced pacing'); Note='Use the extra days wisely.'; NoteText='Ten days give you room for Japan’s classic route without filling every hour. Keep transfer days lighter than normal sightseeing days.'},
        @{N=3; Slug='japan-14-day-itinerary'; Short='14 days in Japan'; Desc='Use this realistic 14-day Japan itinerary for first-time visitors to plan Tokyo, Kyoto, Osaka, day trips, transport, and a comfortable two-week route.'; Meta=@('14-day route','Two-week itinerary','Flexible day trips'); Note='Build depth into two weeks.'; NoteText='A 14-day trip allows more time in each base and selective day trips. Add places because they improve the route, not simply because time is available.'},
        @{N=4; Slug='japan-21-day-itinerary'; Short='21 days in Japan'; Desc='Plan a 21-day Japan itinerary for a first three-week trip, with logical city bases, regional options, transfer days, pacing, and practical route advice.'; Meta=@('21-day route','Three-week first trip','Regional options'); Note='Three weeks still need structure.'; NoteText='Use a few strong bases and add regions in a logical direction. More days should create a calmer trip, not a longer list of hotel changes.'},
        @{N=5; Slug='tokyo-kyoto-osaka-itinerary'; Short='Tokyo, Kyoto & Osaka'; Desc='Learn how to divide your days between Tokyo, Kyoto, and Osaka with sample trip lengths, route order, transport advice, and realistic first-trip pacing.'; Meta=@('Day allocation','Classic Japan route','Tokyo to Osaka'); Note='Divide days by priorities.'; NoteText='Tokyo, Kyoto, and Osaka offer different experiences. Give each city enough time for your interests instead of dividing the trip into equal parts automatically.'}
    )
}
if ($sourceLeaf -eq 'Japan Transportation') {
    $category = 'Japan Transportation'
    $posts = @(
        @{N=1; Slug='how-to-get-around-japan'; Url='how-to-get-around-japan'; Short='Getting around Japan'; Desc='Learn how to get around Japan using trains, subways, buses, taxis, flights, ferries, rental cars, and practical transport planning for a first trip.'; Meta=@('Complete transport guide','Trains, buses & flights','First-time friendly'); Note='Choose transport journey by journey.'; NoteText='Japan does not require one transport solution. Match each journey to distance, luggage, price, convenience, and your actual route.'},
        @{N=2; Slug='japan-train-travel-guide'; Url='japan-train-travel-guide-for-first-time'; Short='Japan train travel'; Desc='Use Japan trains confidently with this first-time guide to JR and private railways, train types, tickets, IC cards, transfers, luggage, and station navigation.'; Meta=@('Train types explained','Tickets & transfers','Station navigation'); Note='Learn the next journey, not the whole network.'; NoteText='Check the railway, line, direction, train type, and platform for the trip in front of you. You do not need to memorize Japan’s rail system.'},
        @{N=3; Slug='japan-rail-pass'; Url='japan-rail-pass-worth'; Short='Is the JR Pass worth it?'; Desc='Find out whether the Japan Rail Pass is worth it by comparing current prices, eligible trains, individual tickets, regional passes, and realistic Japan routes.'; Meta=@('JR Pass value','Route calculations','Pass vs tickets'); Note='Build the route before buying a pass.'; NoteText='The JR Pass saves money only when eligible journeys cost more than the pass. City count and trip length alone do not prove value.'},
        @{N=4; Slug='shinkansen-guide'; Url='shinkansen-guide'; Short='Using the Shinkansen'; Desc='Learn how to use Japan’s Shinkansen, including ticket types, reserved seats, luggage rules, station gates, platforms, boarding, and first-time travel tips.'; Meta=@('Bullet train guide','Tickets & seats','Luggage rules'); Note='Prepare before reaching the platform.'; NoteText='Confirm the station, train name, departure time, car, seat, and luggage plan before boarding. The platform process is then straightforward.'},
        @{N=5; Slug='ic-cards-in-japan'; Url='ic-cards-in-japan'; Short='Suica, PASMO & ICOCA'; Desc='Understand IC cards in Japan, including Suica, PASMO, ICOCA, tourist cards, mobile wallets, recharging, refunds, limitations, and nationwide use.'; Meta=@('Suica, PASMO & ICOCA','Tap-and-go travel','Physical & mobile cards'); Note='One compatible IC card is usually enough.'; NoteText='Use the same card to enter and exit, keep a modest balance, and remember that an IC card is convenient stored value—not an unlimited rail pass.'},
        @{N=6; Slug='luggage-forwarding-in-japan'; Url='luggage-forwarding-in-japan'; Short='Japan luggage forwarding'; Desc='Learn how luggage forwarding in Japan works, including Takkyubin costs, delivery times, hotel and airport shipping, suitcase limits, packing, and tracking.'; Meta=@('Takkyubin explained','Hotel-to-hotel delivery','Luggage-free travel'); Note='Keep one night of essentials with you.'; NoteText='Forwarding is reliable, but delivery is not instant. Carry passports, medication, valuables, electronics, and backup clothing in your day bag.'},
        @{N=7; Slug='narita-vs-haneda-airport'; Url='narita-vs-haneda-airport'; Short='Narita vs Haneda'; Desc='Compare Narita and Haneda airports for Tokyo by location, transport, cost, flight schedules, late arrivals, early departures, luggage, and hotel access.'; Meta=@('Tokyo airport comparison','Transfer time & cost','Choose the better flight'); Note='Compare airport to hotel, not airport to city.'; NoteText='Haneda is closer, but Narita can still be the better choice when its flight schedule, fare, or direct route improves the complete journey.'}
    )
}
if ($sourceLeaf -eq 'Tokyo Hub') {
    $category = 'Tokyo Travel Guide'
    $posts = @(
        @{N=1; Slug='tokyo-travel-guide'; Url='tokyo-travel-guide'; Short='First-time Tokyo guide'; Desc='Plan your first Tokyo trip with practical advice on neighborhoods, sightseeing days, transport, accommodation, food, costs, safety, and seasonal travel.'; Meta=@('First-time Tokyo','Areas & transport','Practical planning'); Note='Plan Tokyo by neighborhood clusters.'; NoteText='Group nearby areas into the same day, leave room for walking and station navigation, and avoid crossing the city repeatedly.'},
        @{N=2; Slug='tokyo-3-day-itinerary'; Url='tokyo-3-day-itinerary'; Short='3 days in Tokyo'; Desc='Follow a realistic Tokyo 3-day itinerary for first-time visitors covering Asakusa, Ueno, Shibuya, Shinjuku, Tsukiji, Ginza, and Tokyo Station.'; Meta=@('3-day itinerary','Eastern & western Tokyo','Realistic pacing'); Note='Three days need a focused route.'; NoteText='Use one geographic cluster per day and treat arrival day separately so Tokyo does not become a rushed checklist.'},
        @{N=3; Slug='tokyo-5-day-itinerary'; Url='tokyo-5-day-itinerary'; Short='5 days in Tokyo'; Desc='Use this balanced Tokyo 5-day itinerary to explore traditional districts, modern neighborhoods, gardens, markets, shopping, and optional day trips.'; Meta=@('5-day itinerary','Balanced city route','Flexible fifth day'); Note='Five days allow depth, not clutter.'; NoteText='Use the extra time to slow down, explore smaller neighborhoods, and keep one day flexible for weather or personal interests.'},
        @{N=4; Slug='where-to-stay-in-tokyo'; Url='where-to-stay-in-tokyo'; Short='Where to stay in Tokyo'; Desc='Compare the best areas to stay in Tokyo for first-time visitors, including Shinjuku, Shibuya, Tokyo Station, Ueno, Asakusa, Ginza, and Shinagawa.'; Meta=@('Best Tokyo areas','Hotel location guide','First-time visitors'); Note='Choose the station before the hotel.'; NoteText='A convenient station and simple route to your priorities usually matter more than choosing the most famous Tokyo neighborhood.'},
        @{N=5; Slug='how-to-get-around-tokyo'; Url='how-to-get-around-tokyo'; Short='Tokyo transport'; Desc='Learn how to get around Tokyo using JR trains, Tokyo Metro, Toei Subway, IC cards, buses, taxis, walking, and practical station-navigation tips.'; Meta=@('JR, Metro & Toei','IC cards & passes','Station navigation'); Note='Use the easiest route, not one operator.'; NoteText='Tokyo transport is simpler when you follow the best journey and allow extra time for transfers, exits, and large stations.'},
        @{N=6; Slug='best-day-trips-from-tokyo'; Url='best-day-trips-from-tokyo'; Short='Tokyo day trips'; Desc='Compare the best day trips from Tokyo for first-time visitors, including Kamakura, Nikko, Hakone, Kawaguchiko, Yokohama, and Kawagoe.'; Meta=@('Six day trips','Fuji, culture & nature','Choose by interest'); Note='Protect enough time for Tokyo first.'; NoteText='Choose one day trip for the experience you want, then check weather, transport, opening hours, and the return journey.'}
    )
}
if ($sourceLeaf -eq 'Kyoto Hub') {
    $category = 'Kyoto Travel Guide'
    $posts = @(
        @{N=1; Slug='kyoto-travel-guide'; Url='kyoto-travel-guide'; Short='First-time Kyoto guide'; Desc='Plan your first Kyoto trip with practical advice on sightseeing areas, transport, accommodation, crowds, food, seasons, etiquette, and daily pacing.'; Meta=@('First-time Kyoto','Areas & transport','Crowd-smart planning'); Note='Plan Kyoto by area and time of day.'; NoteText='Group nearby temples, streets, and neighborhoods together, then use early mornings and flexible evenings to avoid unnecessary crossings and crowds.'},
        @{N=2; Slug='kyoto-2-day-itinerary'; Url='kyoto-2-day-itinerary'; Short='2 days in Kyoto'; Desc='Follow a realistic Kyoto 2-day itinerary covering Fushimi Inari, Kiyomizu-dera, Higashiyama, Gion, Arashiyama, Kinkaku-ji, and downtown Kyoto.'; Meta=@('2-day itinerary','Classic Kyoto route','Efficient sightseeing'); Note='Two Kyoto days need clear priorities.'; NoteText='Use one broad side of Kyoto each day, start the busiest sights early, and remove stops instead of rushing when the schedule slips.'},
        @{N=3; Slug='kyoto-3-day-itinerary'; Url='kyoto-3-day-itinerary'; Short='3 days in Kyoto'; Desc='Use this balanced Kyoto 3-day itinerary for Fushimi Inari, Higashiyama, Gion, Arashiyama, Kinkaku-ji, Nijo Castle, Nishiki Market, and downtown.'; Meta=@('3-day itinerary','Eastern & western Kyoto','Balanced pacing'); Note='Three days allow Kyoto to breathe.'; NoteText='Give eastern Kyoto, Arashiyama, and northern-central Kyoto their own days so travel time and crowds do not dominate the experience.'},
        @{N=4; Slug='where-to-stay-in-kyoto'; Url='where-to-stay-in-kyoto'; Short='Where to stay in Kyoto'; Desc='Compare the best areas to stay in Kyoto for first-time visitors, including Kyoto Station, downtown, Gion, Higashiyama, Shijo-Karasuma, and Arashiyama.'; Meta=@('Best Kyoto areas','Hotel & ryokan guide','First-time visitors'); Note='Choose the base for your real itinerary.'; NoteText='Kyoto Station favors transport, downtown balances dining and access, while Gion and Higashiyama prioritize atmosphere and early sightseeing.'}
    )
}
if ($sourceLeaf -eq 'Osaka Hub') {
    $category = 'Osaka Travel Guide'
    $posts = @(
        @{N=1; Slug='osaka-travel-guide'; Url='osaka-travel-guide'; Short='First-time Osaka guide'; Desc='Plan your first Osaka trip with practical advice on neighborhoods, sightseeing days, transport, accommodation, food, day trips, costs, and seasonal travel.'; Meta=@('First-time Osaka','Areas, food & transport','Practical planning'); Note='Plan Osaka by district and evening.'; NoteText='Group nearby districts into the same day, use trains between areas, and save Osaka’s strongest food and nightlife neighborhoods for the evening.'},
        @{N=2; Slug='osaka-2-day-itinerary'; Url='osaka-2-day-itinerary'; Short='2 days in Osaka'; Desc='Follow a realistic Osaka 2-day itinerary covering Osaka Castle, Nakanoshima, Umeda, Shinsekai, Tennoji, Shinsaibashi, Namba, and Dotonbori.'; Meta=@('2-day itinerary','North & south Osaka','Food and night views'); Note='Give each Osaka day a clear direction.'; NoteText='Use day one for historic and modern northern Osaka, then keep day two focused on Minami, Shinsekai, Tennoji, and Dotonbori after dark.'},
        @{N=3; Slug='osaka-3-day-itinerary'; Url='osaka-3-day-itinerary'; Short='3 days in Osaka'; Desc='Use this balanced Osaka 3-day itinerary for Osaka Castle, Nakanoshima, Umeda, Tennoji, Shinsekai, Namba, Shinsaibashi, and Dotonbori.'; Meta=@('3-day itinerary','Balanced Osaka route','Realistic pacing'); Note='Three days let Osaka breathe.'; NoteText='Separate northern Osaka, southern heritage, and Minami into distinct days so travel time and crowded evenings do not dominate the trip.'},
        @{N=4; Slug='where-to-stay-in-osaka'; Url='where-to-stay-in-osaka'; Short='Where to stay in Osaka'; Desc='Compare the best areas to stay in Osaka for first-time visitors, including Namba, Umeda, Shinsaibashi, Tennoji, Honmachi, and Shin-Osaka.'; Meta=@('Best Osaka areas','Hotel location guide','First-time visitors'); Note='Choose the base for your evenings and onward route.'; NoteText='Namba favors food and nightlife, Umeda favors transport and day trips, while other districts suit quieter stays, value, or specific rail connections.'},
        @{N=5; Slug='kyoto-vs-osaka'; Url='kyoto-vs-osaka'; Short='Kyoto or Osaka base?'; Desc='Compare Kyoto and Osaka as a travel base by sightseeing, nightlife, transport, hotel cost, airports, day trips, luggage, and trip length.'; Meta=@('Kyoto vs Osaka','Best Kansai base','Stay or split nights'); Note='Choose by mornings, evenings, and transport.'; NoteText='Kyoto offers atmosphere and early sightseeing access; Osaka offers nightlife, dining, and stronger regional connections. Splitting nights can be the best answer.'}
    )
}
if ($sourceLeaf -eq 'Iceland Trip Planning') {
    $category = 'Iceland Trip Planning'
    $css = $css -replace 'content:"JP"','content:"IS"'
    $posts = @(
        @{N=1; Slug='iceland-travel-guide'; Url='iceland-travel-guide'; Short='First-time Iceland guide'; Desc='Plan your first Iceland trip with practical advice on regions, road trips, seasons, transport, accommodation, safety, costs, packing, and daily pacing.'; Meta=@('First-time Iceland','Routes, weather & driving','Practical planning'); Note='Build Iceland around conditions, not a checklist.'; NoteText='Choose a realistic region, protect time for weather changes, and check official road and safety information throughout the trip.'},
        @{N=2; Slug='how-to-plan-a-trip-to-iceland'; Url='how-to-plan-a-trip-to-iceland'; Short='Plan an Iceland trip'; Desc='Plan an Iceland trip step by step, from choosing dates, trip length, regions, and transport to booking flights, hotels, activities, insurance, and a flexible route.'; Meta=@('Step-by-step plan','Route before bookings','Weather-ready itinerary'); Note='Choose the route before the reservations.'; NoteText='Start with season, usable days, priority regions, and driving comfort. Flights, hotels, and activities become easier once the route is realistic.'},
        @{N=3; Slug='how-many-days-do-you-need-in-iceland'; Url='how-many-days-do-you-need-in-iceland'; Short='How many days in Iceland?'; Desc='Compare realistic 3-, 5-, 7-, and 10-day Iceland trips, including usable sightseeing days, regional routes, Ring Road timing, driving, and seasonal pacing.'; Meta=@('3, 5, 7 & 10 days','Ring Road timing','Realistic pacing'); Note='Count full usable days, not hotel nights.'; NoteText='Arrival, departure, long drives, weather, and seasonal daylight reduce sightseeing time. Match the route to the days you can genuinely use.'},
        @{N=4; Slug='best-time-to-visit-iceland'; Url='best-time-to-visit-iceland'; Short='Best time for Iceland'; Desc='Find the best time to visit Iceland by comparing weather, daylight, crowds, travel costs, road conditions, Northern Lights, and seasonal activities.'; Meta=@('Weather & daylight','Crowds and costs','Northern Lights timing'); Note='Choose the season for the trip you want.'; NoteText='Daylight, road access, weather, prices, and activities change together. Compare the full seasonal trade-off before choosing dates.'},
        @{N=5; Slug='iceland-trip-cost'; Url='iceland-trip-cost'; Short='Iceland trip cost'; Desc='Estimate an Iceland trip cost for 3, 5, 7, or 10 days with practical budget ranges for flights, accommodation, rental cars, fuel, food, activities, and extras.'; Meta=@('3, 5, 7 & 10 days','Budget ranges','Cost-saving choices'); Note='Let the route shape the Iceland budget.'; NoteText='Accommodation, vehicle choice, fuel, meals, and paid activities create the biggest differences. Keep a separate weather and emergency buffer.'},
        @{N=6; Slug='what-to-pack-for-iceland'; Url='what-to-pack-for-iceland'; Short='Iceland packing list'; Desc='Use this first-time Iceland packing list for waterproof layers, warm clothing, footwear, road-trip essentials, electronics, toiletries, medication, and seasonal gear.'; Meta=@('Year-round layers','Waterproof essentials','Seasonal checklist'); Note='Pack for wind, water, and changing conditions.'; NoteText='A flexible layering system, waterproof outerwear, reliable footwear, and accessible warm clothing matter more than packing many separate outfits.'},
        @{N=7; Slug='iceland-travel-mistakes'; Url='iceland-travel-mistakes'; Short='Iceland travel mistakes'; Desc='Avoid common Iceland travel mistakes involving rushed routes, weather, road safety, rental cars, packing, accommodation, nature, activities, and emergency planning.'; Meta=@('First-trip pitfalls','Road and nature safety','Practical fixes'); Note='Conditions always outrank the itinerary.'; NoteText='Slow down, follow closures and warnings, protect flexibility, and abandon a planned stop when weather, roads, fatigue, or safety require it.'}
    )
}
if ($sourceLeaf -eq 'Iceland Itineraries') {
    $category = 'Iceland Itineraries'
    $css = $css -replace 'content:"JP"','content:"IS"'
    $posts = @(
        @{N=1; Slug='iceland-3-day-itinerary'; Url='iceland-3-day-itinerary'; Short='3 days in Iceland'; Desc='Follow a realistic Iceland 3-day itinerary for first-time visitors, with Reykjavik, the Golden Circle, South Coast highlights, transport, and weather-aware pacing.'; Meta=@('3-day itinerary','Reykjavik & South Iceland','Weather-ready route'); Note='Keep three Iceland days focused.'; NoteText='Choose one compact route, protect time for weather changes, and avoid turning a short visit into a long-distance driving challenge.'},
        @{N=2; Slug='iceland-5-day-itinerary'; Url='iceland-5-day-itinerary'; Short='5 days in Iceland'; Desc='Use this balanced Iceland 5-day itinerary for Reykjavik, the Golden Circle, the South Coast, Vik, waterfalls, beaches, and realistic road-trip pacing.'; Meta=@('5-day itinerary','Golden Circle & South Coast','Balanced road trip'); Note='Five days suit southern Iceland.'; NoteText='Move east at a realistic pace, use logical overnight stops, and leave enough flexibility for weather, roads, and unplanned scenery.'},
        @{N=3; Slug='iceland-7-day-itinerary'; Url='iceland-7-day-itinerary'; Short='One week in Iceland'; Desc='Plan a realistic Iceland 7-day itinerary with Reykjavik, the Golden Circle, South Coast, southeast glacier scenery, accommodation, driving, and flexible pacing.'; Meta=@('7-day itinerary','One week in Iceland','Flexible regional route'); Note='Use one week for depth, not distance.'; NoteText='A strong seven-day trip explores the south and southwest properly instead of forcing a rushed lap around the Ring Road.'},
        @{N=4; Slug='iceland-10-day-itinerary'; Url='iceland-10-day-itinerary'; Short='10-day Ring Road trip'; Desc='Follow a realistic Iceland 10-day itinerary around the Ring Road, including the South Coast, Eastfjords, Myvatn, Akureyri, western Iceland, and Reykjavik.'; Meta=@('10-day itinerary','First Ring Road trip','Day-by-day route'); Note='Ten days make the Ring Road realistic.'; NoteText='Balance daily driving, protect weather flexibility, and treat each overnight stop as part of a continuous route rather than a race around Iceland.'}
    )
}
if ($sourceLeaf -eq 'Driving & Transportation') {
    $category = 'Iceland Driving & Transportation'
    $css = $css -replace 'content:"JP"','content:"IS"'
    $posts = @(
        @{N=1; Slug='how-to-get-around-iceland'; Url='how-to-get-around-iceland'; Short='Getting around Iceland'; Desc='Learn how to get around Iceland by rental car, bus, guided tour, domestic flight, ferry, and airport transfer, with practical advice for first-time visitors.'; Meta=@('Cars, buses & tours','Domestic travel options','First-time transport guide'); Note='Choose transport around your route and season.'; NoteText='A rental car offers flexibility, while tours remove driving pressure. Match the transport plan to weather, distance, daylight, and confidence.'},
        @{N=2; Slug='driving-in-iceland'; Url='driving-in-iceland'; Short='Driving in Iceland'; Desc='Drive in Iceland safely with practical guidance on road rules, weather, wind, gravel, one-lane bridges, winter conditions, closures, and daily route planning.'; Meta=@('Road rules & safety','Weather-ready driving','First-time self-drive'); Note='Conditions always outrank the itinerary.'; NoteText='Check weather, wind, road conditions, and closures before every driving day, then shorten or change the route whenever safety requires it.'},
        @{N=3; Slug='renting-a-car-in-iceland'; Url='renting-a-car-in-iceland'; Short='Renting a car in Iceland'; Desc='Compare 2WD and 4WD rental cars in Iceland, insurance options, costs, fuel, deposits, F-road rules, pickup locations, and common first-time mistakes.'; Meta=@('2WD vs 4WD','Insurance & costs','Rental checklist'); Note='Choose the route before the vehicle.'; NoteText='Season, road type, luggage, group size, and insurance needs should determine the rental car—not appearance or the cheapest headline price.'},
        @{N=4; Slug='iceland-ring-road-guide'; Url='iceland-ring-road-guide'; Short='Iceland Ring Road guide'; Desc='Plan an Iceland Ring Road trip with practical guidance on route direction, major regions, overnight stops, driving time, seasons, safety, and realistic trip length.'; Meta=@('Route 1 guide','Stops & trip length','Weather-aware road trip'); Note='Treat the Ring Road as a journey, not a race.'; NoteText='Balance overnight stops, limit daily driving, protect weather flexibility, and allow each region enough time to feel different from the road between hotels.'}
    )
}

$files = Get-ChildItem -LiteralPath $SourceDir -Filter '*.docx' | Sort-Object Name
if ($files.Count -ne $posts.Count) { throw "Expected $($posts.Count) DOCX files, found $($files.Count)." }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

for ($pi=0; $pi -lt $files.Count; $pi++) {
    $post = $posts[$pi]
    $items = @(Read-Docx $files[$pi].FullName)
    $titleItem = $items | Where-Object { $_.Style -eq 'Heading1' } | Select-Object -First 1
    $title = ($titleItem.Text -replace '\?$', '').Trim()
    $firstPara = $items | Where-Object { $_.Type -eq 'p' -and -not $_.Style } | Select-Object -First 1
    $desc = $post.Desc
    $h2s = @($items | Where-Object { $_.Type -eq 'p' -and $_.Style -eq 'Heading2' })
    $ids = @{}; foreach($h2 in $h2s){ $base=Slug $h2.Text; $id=$base; $k=2; while($ids.ContainsKey($id)){$id="$base-$k";$k++}; $ids[$id]=$h2.Text }
    $idQueue = [Collections.Queue]::new(); foreach($id in $ids.Keys){ $null=$idQueue.Enqueue($id) }
    # Hashtable order is not guaranteed; rebuild IDs in document order.
    $idQueue.Clear(); $used=@{}; foreach($h2 in $h2s){$base=Slug $h2.Text;$id=$base;$k=2;while($used.ContainsKey($id)){$id="$base-$k";$k++};$used[$id]=$true;$idQueue.Enqueue($id)}

    $body = [Text.StringBuilder]::new()
    $faqMode = $false; $faqOpen = $false; $listOpen = $false
    $tokyoMode = $category -in @('Tokyo Travel Guide','Kyoto Travel Guide','Osaka Travel Guide','Iceland Trip Planning','Iceland Itineraries','Iceland Driving & Transportation')
    $titleSkipped = $false
    $tokyoMajorMode = $false
    $tocEntries = [Collections.Generic.List[object]]::new()
    $generatedIds = @{}
    function Close-List { if($script:listOpen){[void]$script:body.Append('</ul>');$script:listOpen=$false} }
    function Close-Faq { if($script:faqOpen){[void]$script:body.Append('</div>');$script:faqOpen=$false} }
    for ($i=0; $i -lt $items.Count; $i++) {
        $it=$items[$i]
        if ($it.Type -eq 'table') {
            Close-List
            [void]$body.Append('<div class="table-wrap"><table>')
            for($ri=0;$ri -lt $it.Rows.Count;$ri++){ $tag=if($ri -eq 0){'th'}else{'td'}; [void]$body.Append('<tr>'); foreach($c in $it.Rows[$ri]){[void]$body.Append("<$tag>$(Html $c)</$tag>")};[void]$body.Append('</tr>') }
            [void]$body.Append('</table></div>'); continue
        }
        if ($it.Style -eq 'Heading1') {
            if (-not $titleSkipped) { $titleSkipped=$true; continue }
            if (-not $tokyoMode) { continue }
            Close-List; Close-Faq
            $base=Slug $it.Text; $id=$base; $k=2; while($generatedIds.ContainsKey($id)){$id="$base-$k";$k++};$generatedIds[$id]=$true
            $faqMode = $it.Text -match '^Frequently Asked Questions'
            $tokyoMajorMode = $true
            [void]$body.Append("<h2 class=`"section-title`" id=`"$id`">$($it.Html)</h2>")
            $tocEntries.Add([pscustomobject]@{Id=$id;Text=$it.Text})
            if($faqMode){[void]$body.Append('<div class="rp-faq-list">');$faqOpen=$true}
            continue
        }
        if ($it.Style -eq 'Heading2') {
            Close-List
            if (-not ($tokyoMode -and $faqMode)) { Close-Faq }
            if ($tokyoMode -and $faqMode) {
                $qBase=Slug $it.Text; $qId=$qBase; $qk=2
                if($generatedIds.ContainsKey($qId)){$qId="$qBase-faq"}
                while($generatedIds.ContainsKey($qId)){$qId="$qBase-faq-$qk";$qk++}
                $generatedIds[$qId]=$true
                [void]$body.Append("<details class=`"rp-faq-item`"><summary class=`"rp-faq-question`" id=`"$qId`"><span>$($it.Html)</span><span aria-hidden=`"true`" class=`"rp-faq-icon`"></span></summary><div class=`"rp-faq-answer`"><div class=`"rp-faq-answer-inner`">")
                $j=$i+1; while($j -lt $items.Count -and $items[$j].Type -eq 'p' -and -not $items[$j].Style){[void]$body.Append("<p>$($items[$j].Html)</p>");$j++};[void]$body.Append('</div></div></details>');$i=$j-1
                continue
            }
            if ($tokyoMode -and $tokyoMajorMode) {
                [void]$body.Append("<h3 class=`"sub-title`">$($it.Html)</h3>")
                continue
            }
            $id=$idQueue.Dequeue()
            $faqMode = $it.Text -match '^Frequently Asked Questions$'
            [void]$body.Append("<h2 class=`"section-title`" id=`"$id`">$($it.Html)</h2>")
            if($tokyoMode){$generatedIds[$id]=$true;$tocEntries.Add([pscustomobject]@{Id=$id;Text=$it.Text})}
            if($faqMode){[void]$body.Append('<div class="rp-faq-list">');$faqOpen=$true}
            continue
        }
        if ($it.Style -eq 'Heading3') {
            Close-List
            if($faqMode){
                $qId=Slug $it.Text
                [void]$body.Append("<details class=`"rp-faq-item`"><summary class=`"rp-faq-question`" id=`"$qId`"><span>$($it.Html)</span><span aria-hidden=`"true`" class=`"rp-faq-icon`"></span></summary><div class=`"rp-faq-answer`"><div class=`"rp-faq-answer-inner`">")
                # FAQ answers are closed when the next heading is reached.
                $j=$i+1; while($j -lt $items.Count -and $items[$j].Type -eq 'p' -and -not $items[$j].Style){[void]$body.Append("<p>$($items[$j].Html)</p>");$j++};[void]$body.Append('</div></div></details>');$i=$j-1
            } else {[void]$body.Append("<h3 class=`"sub-title`">$($it.Html)</h3>")}
            continue
        }
        $plain=$it.Text.Trim()
        $listContext = $it.IsList
        if($listContext){if(-not $listOpen){[void]$body.Append('<ul class="rp-list">');$listOpen=$true};[void]$body.Append("<li><span>&#10003;</span><div>$($it.Html)</div></li>");continue}
        Close-List
        [void]$body.Append("<p>$($it.Html)</p>")
    }
    Close-List; Close-Faq

    $bodyHtml=$body.ToString()
    if ($category -eq 'Japan Transportation') {
        if ($post.Slug -eq 'how-to-get-around-japan') {
            $needle = 'turning transport planning into the hardest part of your Japan trip.</p>'
            $insert = 'turning transport planning into the hardest part of your Japan trip.</p><div class="rp-callout"><div class="callout-icon">&#8596;</div><div><strong>What this guide helps you decide</strong><p>Use this page to choose between trains, subways, buses, taxis, rental cars, domestic flights, ferries, and airport transport. If you already know you will travel by rail and need help with train types, tickets, platforms, transfers, and station exits, use the <a href="https://roamplans.com/japan-train-travel-guide-for-first-time/">Japan Train Travel Guide for First-Time Visitors</a>.</p></div></div>'
            $bodyHtml = $bodyHtml.Replace($needle,$insert)
        }
        if ($post.Slug -eq 'japan-train-travel-guide') {
            $needle = 'because each has its own detailed guide.</p>'
            $insert = 'because each has its own detailed guide.</p><div class="rp-callout"><div class="callout-icon">&#128646;</div><div><strong>This is the train-use guide</strong><p>Use this page when you need to understand JR and private railways, local and rapid trains, tickets, IC-card entry, platforms, transfers, station exits, luggage, and train etiquette. For deciding whether a train, bus, taxi, flight, ferry, or rental car is best for a particular journey, use the broader <a href="https://roamplans.com/how-to-get-around-japan/">Japan Transportation Guide</a>.</p></div></div>'
            $bodyHtml = $bodyHtml.Replace($needle,$insert)
        }
        if ($post.Slug -eq 'japan-rail-pass') {
            $old = '<h2 class="section-title" id="important-october-2026-price-change"><strong>Important October 2026 Price Change</strong></h2><p>There is an important change for travelers purchasing through overseas JR-designated agencies.</p>'
            $new = '<h2 class="section-title" id="important-october-2026-price-change"><strong>Important October 2026 Price Change</strong></h2><div class="rp-callout"><div class="callout-icon">!</div><div><strong>Checked September 8, 2026</strong><p>The change below applies to exchange orders purchased through overseas JR-designated agencies on or after October 1, 2026. It does not mean every purchase channel changes to the same price on that date.</p></div></div>'
            $bodyHtml = $bodyHtml.Replace($old,$new)
            $old = '<p>Because purchase methods and prices can change, always check the official price immediately before paying.</p>'
            $new = '<p>Before paying, confirm both the purchase channel and latest price on the <a href="https://japanrailpass.net/en/purchase/price/" rel="noopener noreferrer">official JAPAN RAIL PASS price page</a>. Do not use an overseas-agency price to calculate an official-online purchase, or assume that a future agency price already applies before its effective date.</p>'
            $bodyHtml = $bodyHtml.Replace($old,$new)
        }
        $internalLinks = @(
            @{Text='Japan Train Travel Guide for First-Time Visitors'; Url='https://roamplans.com/japan-train-travel-guide-for-first-time/'},
            @{Text='How to Get Around Japan: Transportation Guide'; Url='https://roamplans.com/how-to-get-around-japan/'},
            @{Text='Japan Rail Pass: Is the JR Pass Worth It'; Url='https://roamplans.com/japan-rail-pass-worth/'},
            @{Text="Shinkansen Guide: How to Use Japan's Bullet Trains"; Url='https://roamplans.com/shinkansen-guide/'},
            @{Text='IC Cards in Japan: Suica, PASMO and ICOCA Explained'; Url='https://roamplans.com/ic-cards-in-japan/'},
            @{Text='Luggage Forwarding in Japan: How Takkyubin Works'; Url='https://roamplans.com/luggage-forwarding-in-japan/'},
            @{Text='Narita vs Haneda Airport: Which Tokyo Airport Is Better?'; Url='https://roamplans.com/narita-vs-haneda-airport/'},
            @{Text='Japan 7-Day Itinerary for First-Time Visitors'; Url='https://roamplans.com/japan-7-day-itinerary/'},
            @{Text='Japan 10-Day Itinerary for First-Time Visitors'; Url='https://roamplans.com/japan-10-day-itinerary/'},
            @{Text='Japan 14-Day Itinerary for First-Time Visitors'; Url='https://roamplans.com/japan-14-day-itinerary/'},
            @{Text='Japan 21-Day Itinerary: Three-Week First Trip'; Url='https://roamplans.com/japan-21-day-itinerary/'},
            @{Text='Tokyo, Kyoto and Osaka Itinerary: How to Divide Your Days'; Url='https://roamplans.com/tokyo-kyoto-osaka-itinerary/'},
            @{Text='What to Pack for Japan: First-Time Visitor Packing List'; Url='https://roamplans.com/what-to-pack-for-japan/'},
            @{Text='Japan Trip Cost: Budget for 7, 10 and 14 Days'; Url='https://roamplans.com/japan-trip-cost/'},
            @{Text='Japan Travel Mistakes First-Time Visitors Should Avoid'; Url='https://roamplans.com/japan-travel-mistakes/'}
        )
        foreach ($link in $internalLinks) {
            $encoded = Html $link.Text
            if ($bodyHtml -notmatch ('href="' + [regex]::Escape($link.Url) + '"')) {
                $bodyHtml = ([regex]::new([regex]::Escape($encoded))).Replace($bodyHtml,('<a href="' + $link.Url + '">' + $encoded + ' &#8599;</a>'),1)
            }
        }
        $related = switch ($post.Slug) {
            'how-to-get-around-japan' { @(
                @{Url='https://roamplans.com/japan-train-travel-guide-for-first-time/';Text='Learn how to use Japan trains step by step'},
                @{Url='https://roamplans.com/ic-cards-in-japan/';Text='Choose and use a Suica, PASMO, or ICOCA card'},
                @{Url='https://roamplans.com/narita-vs-haneda-airport/';Text='Compare Narita and Haneda airport transfers'}
            ) }
            'japan-train-travel-guide' { @(
                @{Url='https://roamplans.com/shinkansen-guide/';Text='Use the Shinkansen confidently'},
                @{Url='https://roamplans.com/japan-rail-pass-worth/';Text='Calculate whether the JR Pass is worth it'},
                @{Url='https://roamplans.com/luggage-forwarding-in-japan/';Text='Plan luggage forwarding between cities'}
            ) }
            'japan-rail-pass' { @(
                @{Url='https://roamplans.com/shinkansen-guide/';Text='Understand Shinkansen trains, tickets, and seats'},
                @{Url='https://roamplans.com/japan-train-travel-guide-for-first-time/';Text='Read the complete first-time Japan train guide'},
                @{Url='https://roamplans.com/tokyo-kyoto-osaka-itinerary/';Text='Build your Tokyo, Kyoto, and Osaka route before calculating'}
            ) }
            'shinkansen-guide' { @(
                @{Url='https://roamplans.com/japan-rail-pass-worth/';Text='Check whether the JR Pass fits your route'},
                @{Url='https://roamplans.com/luggage-forwarding-in-japan/';Text='Compare forwarding with carrying luggage on the train'},
                @{Url='https://roamplans.com/japan-train-travel-guide-for-first-time/';Text='Learn local trains, transfers, and station navigation'}
            ) }
            'ic-cards-in-japan' { @(
                @{Url='https://roamplans.com/japan-train-travel-guide-for-first-time/';Text='Use IC cards as part of normal train travel'},
                @{Url='https://roamplans.com/shinkansen-guide/';Text='See the separate process for Shinkansen tickets'},
                @{Url='https://roamplans.com/how-to-get-around-japan/';Text='Compare all transportation options in Japan'}
            ) }
            'luggage-forwarding-in-japan' { @(
                @{Url='https://roamplans.com/shinkansen-guide/';Text='Check Shinkansen oversized-luggage rules'},
                @{Url='https://roamplans.com/what-to-pack-for-japan/';Text='Build a lighter first-time Japan packing list'},
                @{Url='https://roamplans.com/japan-10-day-itinerary/';Text='See where forwarding fits a 10-day route'}
            ) }
            'narita-vs-haneda-airport' { @(
                @{Url='https://roamplans.com/how-to-get-around-japan/';Text='Plan transport for the rest of your Japan trip'},
                @{Url='https://roamplans.com/luggage-forwarding-in-japan/';Text='Understand airport luggage delivery'},
                @{Url='https://roamplans.com/tokyo-kyoto-osaka-itinerary/';Text='Plan an open-jaw Tokyo, Kyoto, and Osaka route'}
            ) }
        }
        $relatedItems = ($related | ForEach-Object { '<li><span>&#8599;</span><div><a href="' + $_.Url + '">' + (Html $_.Text) + '</a></div></li>' }) -join ''
        $relatedBlock = '<div class="rp-callout"><div class="callout-icon">+</div><div><strong>Related Japan transport guides</strong><ul class="rp-list">' + $relatedItems + '</ul></div></div>'
        $faqHeading = '<h2 class="section-title" id="frequently-asked-questions">'
        $bodyHtml = $bodyHtml.Replace($faqHeading,($relatedBlock + $faqHeading))
    }
    if ($category -eq 'Tokyo Travel Guide') {
        $related = switch ($post.Slug) {
            'tokyo-travel-guide' { @(
                @{Url='https://roamplans.com/tokyo-3-day-itinerary/';Text='Follow the focused three-day Tokyo route'},
                @{Url='https://roamplans.com/where-to-stay-in-tokyo/';Text='Choose the best Tokyo area for your stay'},
                @{Url='https://roamplans.com/how-to-get-around-tokyo/';Text='Understand Tokyo trains, subways, and IC cards'}
            ) }
            'tokyo-3-day-itinerary' { @(
                @{Url='https://roamplans.com/tokyo-5-day-itinerary/';Text='Compare the more relaxed five-day Tokyo route'},
                @{Url='https://roamplans.com/where-to-stay-in-tokyo/';Text='Choose a convenient base for this itinerary'},
                @{Url='https://roamplans.com/how-to-get-around-tokyo/';Text='Plan trains and subway journeys between each area'}
            ) }
            'tokyo-5-day-itinerary' { @(
                @{Url='https://roamplans.com/tokyo-3-day-itinerary/';Text='See the shorter essential Tokyo itinerary'},
                @{Url='https://roamplans.com/best-day-trips-from-tokyo/';Text='Choose an optional day trip from Tokyo'},
                @{Url='https://roamplans.com/where-to-stay-in-tokyo/';Text='Find the best area for a five-night stay'}
            ) }
            'where-to-stay-in-tokyo' { @(
                @{Url='https://roamplans.com/tokyo-3-day-itinerary/';Text='Match your hotel base to a three-day itinerary'},
                @{Url='https://roamplans.com/tokyo-5-day-itinerary/';Text='Choose a base for a longer Tokyo stay'},
                @{Url='https://roamplans.com/narita-vs-haneda-airport/';Text='Compare airport access before booking'}
            ) }
            'how-to-get-around-tokyo' { @(
                @{Url='https://roamplans.com/ic-cards-in-japan/';Text='Learn how Suica, PASMO, and ICOCA work'},
                @{Url='https://roamplans.com/japan-train-travel-guide-for-first-time/';Text='Read the complete Japan train guide'},
                @{Url='https://roamplans.com/narita-vs-haneda-airport/';Text='Plan your airport-to-Tokyo transfer'}
            ) }
            'best-day-trips-from-tokyo' { @(
                @{Url='https://roamplans.com/tokyo-5-day-itinerary/';Text='Decide where a day trip fits into five Tokyo days'},
                @{Url='https://roamplans.com/how-to-get-around-tokyo/';Text='Prepare Tokyo transport before departure'},
                @{Url='https://roamplans.com/japan-rail-pass-worth/';Text='Check whether your longer journeys justify a rail pass'}
            ) }
        }
        $relatedItems = ($related | ForEach-Object { '<li><span>&#8599;</span><div><a href="' + $_.Url + '">' + (Html $_.Text) + '</a></div></li>' }) -join ''
        $relatedBlock = '<div class="rp-callout"><div class="callout-icon">+</div><div><strong>Related Tokyo planning guides</strong><ul class="rp-list">' + $relatedItems + '</ul></div></div>'
        $faqPattern = '<h2 class="section-title" id="frequently-asked-questions(?:-about-visiting-tokyo)?">'
        $bodyHtml = ([regex]$faqPattern).Replace($bodyHtml,($relatedBlock + '$0'),1)
    }
    if ($category -eq 'Kyoto Travel Guide') {
        $related = switch ($post.Slug) {
            'kyoto-travel-guide' { @(
                @{Url='https://roamplans.com/kyoto-2-day-itinerary/';Text='Follow the focused two-day Kyoto route'},
                @{Url='https://roamplans.com/where-to-stay-in-kyoto/';Text='Choose the best Kyoto area for your stay'},
                @{Url='https://roamplans.com/kyoto-3-day-itinerary/';Text='Plan a more balanced three-day visit'}
            ) }
            'kyoto-2-day-itinerary' { @(
                @{Url='https://roamplans.com/kyoto-3-day-itinerary/';Text='Compare the more relaxed three-day Kyoto route'},
                @{Url='https://roamplans.com/where-to-stay-in-kyoto/';Text='Choose a convenient base for two days'},
                @{Url='https://roamplans.com/japan-train-travel-guide-for-first-time/';Text='Prepare for rail travel to and from Kyoto'}
            ) }
            'kyoto-3-day-itinerary' { @(
                @{Url='https://roamplans.com/kyoto-2-day-itinerary/';Text='See the shorter essential Kyoto route'},
                @{Url='https://roamplans.com/where-to-stay-in-kyoto/';Text='Match your hotel area to this itinerary'},
                @{Url='https://roamplans.com/tokyo-kyoto-osaka-itinerary/';Text='Fit Kyoto into the classic Japan route'}
            ) }
            'where-to-stay-in-kyoto' { @(
                @{Url='https://roamplans.com/kyoto-2-day-itinerary/';Text='Match your base to a two-day itinerary'},
                @{Url='https://roamplans.com/kyoto-3-day-itinerary/';Text='Choose a base for three Kyoto days'},
                @{Url='https://roamplans.com/luggage-forwarding-in-japan/';Text='Plan luggage forwarding to your Kyoto hotel'}
            ) }
        }
        $relatedItems = ($related | ForEach-Object { '<li><span>&#8599;</span><div><a href="' + $_.Url + '">' + (Html $_.Text) + '</a></div></li>' }) -join ''
        $relatedBlock = '<div class="rp-callout"><div class="callout-icon">+</div><div><strong>Related Kyoto planning guides</strong><ul class="rp-list">' + $relatedItems + '</ul></div></div>'
        $faqPattern = '<h2 class="section-title" id="frequently-asked-questions">'
        $bodyHtml = ([regex]$faqPattern).Replace($bodyHtml,($relatedBlock + '$0'),1)
    }
    if ($category -eq 'Osaka Travel Guide') {
        if ($post.Slug -eq 'osaka-2-day-itinerary') {
            $day2Pattern = '(<strong> all in one exhausting first </strong><strong>day\.</strong>).*?</p>'
            $day2Heading = '$1</p><h2 class="section-title" id="day-2-kuromon-market-namba-shinsekai-tennoji-shinsaibashi-and-dotonbori">Day 2: Kuromon Market, Namba, Shinsekai, Tennoji, Shinsaibashi and Dotonbori</h2>'
            $bodyHtml = ([regex]$day2Pattern).Replace($bodyHtml,$day2Heading,1)
            if(-not $generatedIds.ContainsKey('day-2-kuromon-market-namba-shinsekai-tennoji-shinsaibashi-and-dotonbori')){$tocEntries.Add([pscustomobject]@{Id='day-2-kuromon-market-namba-shinsekai-tennoji-shinsaibashi-and-dotonbori';Text='Day 2: Kuromon Market, Namba, Shinsekai, Tennoji, Shinsaibashi and Dotonbori'})}
        }
        $related = switch ($post.Slug) {
            'osaka-travel-guide' { @(
                @{Url='https://roamplans.com/osaka-2-day-itinerary/';Text='Follow the focused two-day Osaka route'},
                @{Url='https://roamplans.com/where-to-stay-in-osaka/';Text='Choose the best Osaka area for your stay'},
                @{Url='https://roamplans.com/kyoto-vs-osaka/';Text='Compare Kyoto and Osaka as your Kansai base'}
            ) }
            'osaka-2-day-itinerary' { @(
                @{Url='https://roamplans.com/osaka-3-day-itinerary/';Text='Compare the more relaxed three-day Osaka route'},
                @{Url='https://roamplans.com/where-to-stay-in-osaka/';Text='Choose a convenient base for two days'},
                @{Url='https://roamplans.com/kyoto-vs-osaka/';Text='Decide whether to stay in Osaka or Kyoto'}
            ) }
            'osaka-3-day-itinerary' { @(
                @{Url='https://roamplans.com/osaka-2-day-itinerary/';Text='See the shorter essential Osaka route'},
                @{Url='https://roamplans.com/where-to-stay-in-osaka/';Text='Match your hotel area to this itinerary'},
                @{Url='https://roamplans.com/kyoto-vs-osaka/';Text='Compare Osaka with a Kyoto base'}
            ) }
            'where-to-stay-in-osaka' { @(
                @{Url='https://roamplans.com/osaka-2-day-itinerary/';Text='Match your base to a two-day itinerary'},
                @{Url='https://roamplans.com/osaka-3-day-itinerary/';Text='Choose a base for three Osaka days'},
                @{Url='https://roamplans.com/kyoto-vs-osaka/';Text='Compare staying in Osaka with Kyoto'}
            ) }
            'kyoto-vs-osaka' { @(
                @{Url='https://roamplans.com/where-to-stay-in-kyoto/';Text='Compare the best areas to stay in Kyoto'},
                @{Url='https://roamplans.com/where-to-stay-in-osaka/';Text='Compare the best areas to stay in Osaka'},
                @{Url='https://roamplans.com/tokyo-kyoto-osaka-itinerary/';Text='Fit both cities into the classic Japan route'}
            ) }
        }
        $relatedItems = ($related | ForEach-Object { '<li><span>&#8599;</span><div><a href="' + $_.Url + '">' + (Html $_.Text) + '</a></div></li>' }) -join ''
        $relatedBlock = '<div class="rp-callout"><div class="callout-icon">+</div><div><strong>Related Osaka planning guides</strong><ul class="rp-list">' + $relatedItems + '</ul></div></div>'
        $faqPattern = '<h2 class="section-title" id="frequently-asked-questions(?:-about-osaka)?">'
        $bodyHtml = ([regex]$faqPattern).Replace($bodyHtml,($relatedBlock + '$0'),1)
    }
    if ($category -in @('Iceland Trip Planning','Iceland Itineraries','Iceland Driving & Transportation')) {
        # Some source Word files apply bold to entire normal paragraphs. Restore
        # those paragraphs to the locked layout's body weight while preserving
        # genuine inline emphasis elsewhere.
        $bodyHtml = [regex]::Replace($bodyHtml, '<p>((?:\s*<strong>.*?</strong>\s*)+)</p>', {
            param($match)
            '<p>' + ($match.Groups[1].Value -replace '</?strong>', '') + '</p>'
        })

        $related = switch ($post.Slug) {
            'iceland-3-day-itinerary' { @(
                @{Url='https://roamplans.com/iceland-5-day-itinerary/';Text='Compare the more flexible five-day Iceland route'},
                @{Url='https://roamplans.com/iceland-travel-guide/';Text='Review the complete first-time Iceland guide'},
                @{Url='https://roamplans.com/best-time-to-visit-iceland/';Text='Match this short route to season and daylight'}
            ) }
            'iceland-5-day-itinerary' { @(
                @{Url='https://roamplans.com/iceland-3-day-itinerary/';Text='See the compact three-day Iceland route'},
                @{Url='https://roamplans.com/iceland-7-day-itinerary/';Text='Expand the trip to one full week'},
                @{Url='https://roamplans.com/iceland-trip-cost/';Text='Estimate the cost of a five-day Iceland trip'}
            ) }
            'iceland-7-day-itinerary' { @(
                @{Url='https://roamplans.com/iceland-5-day-itinerary/';Text='Compare the focused five-day South Iceland route'},
                @{Url='https://roamplans.com/iceland-10-day-itinerary/';Text='Compare the complete ten-day Ring Road route'},
                @{Url='https://roamplans.com/how-many-days-do-you-need-in-iceland/';Text='Choose the right Iceland trip length'}
            ) }
            'iceland-10-day-itinerary' { @(
                @{Url='https://roamplans.com/iceland-7-day-itinerary/';Text='Compare a slower one-week regional route'},
                @{Url='https://roamplans.com/iceland-travel-guide/';Text='Review the complete first-time Iceland guide'},
                @{Url='https://roamplans.com/iceland-travel-mistakes/';Text='Avoid common Ring Road planning mistakes'}
            ) }
            'how-to-get-around-iceland' { @(
                @{Url='https://roamplans.com/driving-in-iceland/';Text='Prepare for Iceland road rules, weather, and safety'},
                @{Url='https://roamplans.com/renting-a-car-in-iceland/';Text='Choose the right rental vehicle and insurance'},
                @{Url='https://roamplans.com/iceland-ring-road-guide/';Text='Plan transport for a complete Ring Road trip'}
            ) }
            'driving-in-iceland' { @(
                @{Url='https://roamplans.com/renting-a-car-in-iceland/';Text='Compare 2WD, 4WD, insurance, and rental costs'},
                @{Url='https://roamplans.com/how-to-get-around-iceland/';Text='Compare driving with tours, buses, and flights'},
                @{Url='https://roamplans.com/iceland-travel-mistakes/';Text='Avoid common Iceland driving and safety mistakes'}
            ) }
            'renting-a-car-in-iceland' { @(
                @{Url='https://roamplans.com/driving-in-iceland/';Text='Learn Iceland road rules and safe-driving practices'},
                @{Url='https://roamplans.com/iceland-ring-road-guide/';Text='Match your rental car to the Ring Road route'},
                @{Url='https://roamplans.com/iceland-trip-cost/';Text='Add vehicle, insurance, and fuel to your budget'}
            ) }
            'iceland-ring-road-guide' { @(
                @{Url='https://roamplans.com/iceland-10-day-itinerary/';Text='Follow a realistic ten-day Ring Road itinerary'},
                @{Url='https://roamplans.com/driving-in-iceland/';Text='Prepare for road conditions, weather, and safety'},
                @{Url='https://roamplans.com/renting-a-car-in-iceland/';Text='Choose the right vehicle for your route and season'}
            ) }
            'iceland-travel-guide' { @(
                @{Url='https://roamplans.com/how-to-plan-a-trip-to-iceland/';Text='Build your Iceland trip step by step'},
                @{Url='https://roamplans.com/how-many-days-do-you-need-in-iceland/';Text='Choose a realistic Iceland trip length'},
                @{Url='https://roamplans.com/best-time-to-visit-iceland/';Text='Compare Iceland seasons, weather, and daylight'}
            ) }
            'how-to-plan-a-trip-to-iceland' { @(
                @{Url='https://roamplans.com/iceland-travel-guide/';Text='Start with the complete first-time Iceland guide'},
                @{Url='https://roamplans.com/iceland-trip-cost/';Text='Estimate the budget before booking'},
                @{Url='https://roamplans.com/iceland-travel-mistakes/';Text='Avoid the most common planning and driving mistakes'}
            ) }
            'how-many-days-do-you-need-in-iceland' { @(
                @{Url='https://roamplans.com/how-to-plan-a-trip-to-iceland/';Text='Turn your available days into a realistic route'},
                @{Url='https://roamplans.com/best-time-to-visit-iceland/';Text='Adjust the itinerary for season and daylight'},
                @{Url='https://roamplans.com/iceland-trip-cost/';Text='Compare costs for 3, 5, 7, and 10 days'}
            ) }
            'best-time-to-visit-iceland' { @(
                @{Url='https://roamplans.com/how-many-days-do-you-need-in-iceland/';Text='Match the season to your available days'},
                @{Url='https://roamplans.com/what-to-pack-for-iceland/';Text='Pack for Iceland weather in every season'},
                @{Url='https://roamplans.com/iceland-travel-mistakes/';Text='Avoid weather and road-planning mistakes'}
            ) }
            'iceland-trip-cost' { @(
                @{Url='https://roamplans.com/how-to-plan-a-trip-to-iceland/';Text='Build the route before finalizing the budget'},
                @{Url='https://roamplans.com/how-many-days-do-you-need-in-iceland/';Text='Choose the trip length your budget can support'},
                @{Url='https://roamplans.com/what-to-pack-for-iceland/';Text='Avoid unnecessary gear purchases with a focused list'}
            ) }
            'what-to-pack-for-iceland' { @(
                @{Url='https://roamplans.com/best-time-to-visit-iceland/';Text='Check the seasonal conditions behind the packing list'},
                @{Url='https://roamplans.com/how-to-plan-a-trip-to-iceland/';Text='Connect packing decisions to your route and activities'},
                @{Url='https://roamplans.com/iceland-travel-mistakes/';Text='Avoid common Iceland clothing and footwear mistakes'}
            ) }
            'iceland-travel-mistakes' { @(
                @{Url='https://roamplans.com/iceland-travel-guide/';Text='Review the complete first-time Iceland guide'},
                @{Url='https://roamplans.com/how-to-plan-a-trip-to-iceland/';Text='Use the step-by-step planning workflow'},
                @{Url='https://roamplans.com/what-to-pack-for-iceland/';Text='Prepare waterproof layers and road-trip essentials'}
            ) }
        }
        $relatedItems = ($related | ForEach-Object { '<li><span>&#8599;</span><div><a href="' + $_.Url + '">' + (Html $_.Text) + '</a></div></li>' }) -join ''
        $relatedLabel = if ($category -eq 'Iceland Itineraries') { 'Related Iceland itinerary guides' } elseif ($category -eq 'Iceland Driving & Transportation') { 'Related Iceland driving and transport guides' } else { 'Related Iceland planning guides' }
        $relatedBlock = '<div class="rp-callout"><div class="callout-icon">+</div><div><strong>' + $relatedLabel + '</strong><ul class="rp-list">' + $relatedItems + '</ul></div></div>'
        $faqPattern = '<h2 class="section-title" id="frequently-asked-questions">'
        $bodyHtml = ([regex]$faqPattern).Replace($bodyHtml,($relatedBlock + '$0'),1)
    }
    if ($category -in @('Japan Itineraries','Japan Transportation','Tokyo Travel Guide','Kyoto Travel Guide','Osaka Travel Guide','Iceland Trip Planning','Iceland Itineraries','Iceland Driving & Transportation')) {
        $imagePath = "images/$($post.Slug).png"
        $imageAlt = Html ($title + ' visual overview')
        $imageCaption = Html ($title + ' - a visual overview of the key decisions covered in this guide.')
        $leadImage = "<figure class=`"rp-post-image`"><img src=`"$imagePath`" alt=`"$imageAlt`" width=`"1536`" height=`"1024`" loading=`"eager`" decoding=`"async`"><figcaption>$imageCaption</figcaption></figure>"
        $bodyHtml = ([regex]'<h2 class="section-title" id="[^"]+">.*?</h2>').Replace($bodyHtml, ('$0' + $leadImage), 1)
        $placements = switch ($post.Slug) {
            'japan-7-day-itinerary' {
                @(
                    @{After='day-2-western-tokyo-meiji-jingu-harajuku-shibuya-and-shinjuku'; File='japan-7-day-tokyo.png'; Alt='Meiji Jingu, Harajuku and Shibuya during a seven-day Japan itinerary'; Caption='Western Tokyo combines the calm approach to Meiji Jingu with Harajuku and Shibuya energy.'},
                    @{After='day-5-fushimi-inari-kiyomizudera-and-higashiyama'; File='japan-7-day-kyoto.png'; Alt='Fushimi Inari and Higashiyama during a seven-day Japan itinerary'; Caption='Fushimi Inari and Higashiyama work well together when the Kyoto day starts early.'}
                )
            }
            'japan-10-day-itinerary' {
                @(
                    @{After='day-2-asakusa-ueno-and-eastern-tokyo'; File='japan-10-day-tokyo.png'; Alt='Asakusa, Ueno and eastern Tokyo itinerary'; Caption='Asakusa and Ueno create a practical eastern Tokyo sightseeing day.'},
                    @{After='day-7-arashiyama-and-a-second-side-of-kyoto'; File='japan-10-day-kyoto-osaka.png'; Alt='Arashiyama and Osaka during a ten-day Japan itinerary'; Caption='The longer route leaves room for Kyoto scenery before the trip moves toward Osaka.'}
                )
            }
            'japan-14-day-itinerary' {
                @(
                    @{After='day-8-nara-day-trip-from-kyoto'; File='japan-14-day-nara.png'; Alt='Nara Park day trip during a fourteen-day Japan itinerary'; Caption='A two-week itinerary has enough space for a relaxed Nara day trip from Kyoto.'},
                    @{After='day-10-miyajima-day-trip'; File='japan-14-day-hiroshima-miyajima.png'; Alt='Miyajima day trip from Hiroshima during a Japan itinerary'; Caption='Miyajima adds a distinctive waterfront day to the Hiroshima portion of the route.'}
                )
            }
            'japan-21-day-itinerary' {
                @(
                    @{After='day-7-full-kanazawa-day'; File='japan-21-day-kanazawa.png'; Alt='Kenrokuen and traditional Kanazawa streets during a Japan itinerary'; Caption='Kanazawa rewards a full day with gardens, historic districts, and a slower pace.'},
                    @{After='day-9-full-takayama-day'; File='japan-21-day-takayama.png'; Alt='Takayama and Shirakawa-go mountain region during a three-week Japan itinerary'; Caption='Takayama and the mountain region give the three-week route a different character from the major cities.'}
                )
            }
            'tokyo-kyoto-osaka-itinerary' {
                @(
                    @{After='how-many-days-do-you-need-in-kyoto'; File='tokyo-kyoto-days.png'; Alt='Tokyo and Kyoto city comparison for planning itinerary days'; Caption='Tokyo and Kyoto need different pacing, so equal day counts are not always the best split.'},
                    @{After='how-many-days-do-you-need-in-osaka'; File='osaka-itinerary-days.png'; Alt='Osaka evening itinerary around the Dotonbori canal'; Caption='Osaka can work as a focused day, an evening base, or a longer food-and-neighborhood stay.'}
                )
            }
            'how-to-get-around-japan' {
                @(
                    @{After='trains-are-the-main-transport-for-most-visitors'; File='japan-transport-trains.png'; Alt='Japan trains and urban public transportation for first-time visitors'; Caption='Rail is the backbone of most first-time Japan routes, with local transport filling the gaps.'},
                    @{After='airport-transportation'; File='japan-airport-transport-options.png'; Alt='Japan airport train bus and taxi transportation options'; Caption='Airport transport should be chosen around the hotel location, arrival time, luggage, and total cost.'}
                )
            }
            'japan-train-travel-guide' {
                @(
                    @{After='how-to-use-a-train-station-in-japan'; File='japan-train-station-navigation.png'; Alt='Traveler navigating signs and platforms inside a Japanese train station'; Caption='Clear bilingual signs and departure boards make stations manageable when you follow the journey step by step.'},
                    @{After='train-etiquette-for-first-time-visitors'; File='japan-train-etiquette.png'; Alt='Orderly passengers following train etiquette in Japan'; Caption='Simple habits around queues, doors, luggage, noise, and priority seats make train travel smoother.'}
                )
            }
            'japan-rail-pass' {
                @(
                    @{After='the-biggest-jr-pass-question'; File='jr-pass-cost-calculation.png'; Alt='Japan Rail Pass price compared with individual train ticket costs'; Caption='The useful comparison is the pass price against eligible journeys in the finished itinerary.'},
                    @{After='regional-pass-vs-nationwide-jr-pass'; File='regional-pass-vs-jr-pass.png'; Alt='Regional Japan rail pass compared with nationwide JR Pass'; Caption='A regional pass can fit a concentrated route better than nationwide coverage.'}
                )
            }
            'shinkansen-guide' {
                @(
                    @{After='how-to-board-the-shinkansen-step-by-step'; File='shinkansen-boarding-platform.png'; Alt='Shinkansen platform signs car numbers and boarding queue'; Caption='Match the departure board, platform, train, car number, and marked boarding position before the train arrives.'},
                    @{After='shinkansen-luggage-rules-for-first-time-visitors'; File='shinkansen-luggage-rules.png'; Alt='Suitcase storage and oversized baggage seating on a Shinkansen'; Caption='Measure large luggage early and reserve the appropriate baggage space when the route requires it.'}
                )
            }
            'ic-cards-in-japan' {
                @(
                    @{After='suica-vs-pasmo-vs-icoca-which-is-best'; File='suica-pasmo-icoca-comparison.png'; Alt='Suica PASMO and ICOCA transit cards compared'; Caption='The major compatible cards work similarly for most visitors; availability and arrival region often decide the easiest choice.'},
                    @{After='how-to-use-an-ic-card-at-train-gates'; File='japan-ic-card-ticket-gate.png'; Alt='IC card tapping at a Japanese train station ticket gate'; Caption='Tap the same card when entering and leaving, and keep enough balance for the fare.'}
                )
            }
            'luggage-forwarding-in-japan' {
                @(
                    @{After='how-luggage-forwarding-works-step-by-step'; File='takkyubin-hotel-forwarding.png'; Alt='Hotel staff preparing a suitcase for Takkyubin luggage forwarding'; Caption='Hotel-to-hotel forwarding is easiest when the receiving property accepts luggage and the waybill details match the reservation.'},
                    @{After='sending-luggage-to-an-airport'; File='japan-airport-luggage-forwarding.png'; Alt='Suitcases being forwarded to a Japanese airport service counter'; Caption='Airport delivery needs a larger time buffer and the correct airport terminal information.'}
                )
            }
            'narita-vs-haneda-airport' {
                @(
                    @{After='haneda-vs-narita-do-not-compare-airport-train-time-alone'; File='narita-haneda-to-tokyo-map.png'; Alt='Narita and Haneda airport transport routes to central Tokyo'; Caption='Compare the complete airport-to-hotel journey, including transfers, walking, luggage, and arrival time.'},
                    @{After='narita-express-vs-skyliner-which-is-better'; File='narita-express-vs-skyliner.png'; Alt='Narita Express and Keisei Skyliner airport trains compared'; Caption='The better Narita train depends mainly on the Tokyo neighborhood and transfer required.'}
                )
            }
            'tokyo-travel-guide' {
                @(
                    @{After='understand-tokyo-as-neighborhood-clusters'; File='tokyo-neighborhood-clusters.png'; Alt='Tokyo neighborhood clusters showing western eastern and central sightseeing areas'; Caption='Planning Tokyo by nearby neighborhood clusters reduces unnecessary travel and keeps each day manageable.'},
                    @{After='which-airport-should-you-use-for-tokyo'; File='tokyo-airport-arrival.png'; Alt='First-time visitor arriving in Tokyo and comparing Haneda and Narita transport'; Caption='Choose the airport and transfer by comparing the complete journey to your hotel.'}
                )
            }
            'tokyo-3-day-itinerary' {
                @(
                    @{After='day-2-meiji-jingu-harajuku-shibuya-and-shinjuku'; File='tokyo-3-day-western-tokyo.png'; Alt='Meiji Jingu Harajuku Shibuya and Shinjuku on day two in Tokyo'; Caption='Day two groups western Tokyo into one logical route from Meiji Jingu to Shinjuku.'},
                    @{After='day-3-tsukiji-ginza-tokyo-station-and-marunouchi'; File='tokyo-3-day-central-tokyo.png'; Alt='Tsukiji Ginza Tokyo Station and Marunouchi on a three-day Tokyo itinerary'; Caption='The final day keeps central Tokyo compact while leaving flexible time before departure.'}
                )
            }
            'tokyo-5-day-itinerary' {
                @(
                    @{After='day-2-ueno-yanaka-nezu-and-akihabara'; File='tokyo-5-day-ueno-yanaka.png'; Alt='Ueno Yanaka Nezu and Akihabara route in a five-day Tokyo itinerary'; Caption='Five days create room for quieter historic neighborhoods as well as major Tokyo districts.'},
                    @{After='day-4-shinjuku-gyoen-shinjuku-and-evening-tokyo'; File='tokyo-5-day-shinjuku.png'; Alt='Shinjuku Gyoen and Shinjuku evening during a five-day Tokyo itinerary'; Caption='A full Shinjuku day balances garden time, city views, shopping, and evening atmosphere.'}
                )
            }
            'where-to-stay-in-tokyo' {
                @(
                    @{After='shinjuku-best-all-round-base-for-many-first-time-visitors'; File='where-to-stay-tokyo-shinjuku-shibuya.png'; Alt='Shinjuku and Shibuya accommodation areas compared for first-time Tokyo visitors'; Caption='Shinjuku and Shibuya are convenient modern bases, but their atmosphere and transport strengths differ.'},
                    @{After='ueno-best-for-value-narita-access-and-eastern-tokyo'; File='where-to-stay-tokyo-ueno-asakusa.png'; Alt='Ueno and Asakusa hotel areas in eastern Tokyo'; Caption='Ueno and Asakusa can offer useful value, eastern Tokyo access, and a different evening atmosphere.'}
                )
            }
            'how-to-get-around-tokyo' {
                @(
                    @{After='jr-vs-tokyo-metro-vs-toei-which-should-you-use'; File='tokyo-jr-metro-toei.png'; Alt='JR Tokyo Metro and Toei Subway transport systems compared'; Caption='Use JR, Tokyo Metro, or Toei according to the easiest route rather than committing to one operator.'},
                    @{After='station-exits-can-save-or-waste-time'; File='tokyo-station-exits.png'; Alt='Traveler checking numbered exits inside a large Tokyo station'; Caption='The correct station exit can save a long walk and make large Tokyo stations much easier.'}
                )
            }
            'best-day-trips-from-tokyo' {
                @(
                    @{After='3-hakone-best-for-onsen-lake-scenery-and-a-slower-mountain-day'; File='tokyo-day-trips-hakone-kawaguchiko.png'; Alt='Hakone lake scenery and Mount Fuji view near Kawaguchiko'; Caption='Hakone and Kawaguchiko offer different mountain experiences, with weather playing a major role.'},
                    @{After='5-yokohama-best-easy-urban-day-trip'; File='tokyo-day-trips-yokohama-kawagoe.png'; Alt='Yokohama waterfront and traditional Kawagoe streets as Tokyo day trips'; Caption='Yokohama and Kawagoe are easier alternatives when you want a shorter urban or traditional-town escape.'}
                )
            }
            'kyoto-travel-guide' {
                @(
                    @{After='eastern-kyoto-best-for-the-classic-first-time-experience'; File='kyoto-eastern-higashiyama.png'; Alt='Eastern Kyoto temples and preserved Higashiyama streets for first-time visitors'; Caption='Eastern Kyoto combines major temples, preserved streets, and Gion into one classic first-time area.'},
                    @{After='western-kyoto-arashiyama'; File='kyoto-western-arashiyama.png'; Alt='Arashiyama bamboo and riverside scenery in western Kyoto'; Caption='Arashiyama works best as its own western Kyoto outing with enough time beyond the busiest bamboo path.'}
                )
            }
            'kyoto-2-day-itinerary' {
                @(
                    @{After='day-1-fushimi-inari-kiyomizu-dera-higashiyama-and-gion'; File='kyoto-2-day-fushimi-higashiyama.png'; Alt='Fushimi Inari Kiyomizu-dera Higashiyama and Gion during a two-day Kyoto itinerary'; Caption='Day one moves through southern and eastern Kyoto before finishing around Gion.'},
                    @{After='day-2-arashiyama-kinkaku-ji-nishiki-market-and-downtown-kyoto'; File='kyoto-2-day-arashiyama-downtown.png'; Alt='Arashiyama Kinkaku-ji Nishiki Market and downtown Kyoto itinerary'; Caption='Day two connects western and northern Kyoto with a flexible downtown finish.'}
                )
            }
            'kyoto-3-day-itinerary' {
                @(
                    @{After='day-2-a-full-day-in-arashiyama-and-saga'; File='kyoto-3-day-arashiyama-saga.png'; Alt='Arashiyama and quiet Saga neighborhoods during a three-day Kyoto itinerary'; Caption='A full western Kyoto day leaves room for gardens, riverside scenery, and quieter Saga streets.'},
                    @{After='day-3-kinkaku-ji-ryoan-ji-nijo-castle-nishiki-market-and-downtown-kyoto'; File='kyoto-3-day-northern-central.png'; Alt='Kinkaku-ji Ryoan-ji Nijo Castle and Nishiki Market on day three in Kyoto'; Caption='The third day links northern temples with Kyoto history, food, and downtown streets.'}
                )
            }
            'where-to-stay-in-kyoto' {
                @(
                    @{After='kyoto-station-best-for-transport-short-stays-and-first-time-logistics'; File='where-to-stay-kyoto-station-downtown.png'; Alt='Kyoto Station and downtown Kyoto hotel areas compared'; Caption='Kyoto Station favors transport and short stays, while downtown offers a broader dining and evening base.'},
                    @{After='gion-and-higashiyama-best-for-traditional-kyoto-atmosphere'; File='where-to-stay-kyoto-gion-higashiyama.png'; Alt='Traditional accommodation streets in Gion and Higashiyama Kyoto'; Caption='Gion and Higashiyama prioritize traditional atmosphere and early access to eastern Kyoto.'}
                )
            }
            'osaka-travel-guide' {
                @(
                    @{After='minami-best-for-the-classic-first-time-osaka-experience'; File='osaka-minami-street-food.png'; Alt='Takoyaki cooking in a lively Minami street-food lane at night'; Caption='Minami is at its best when street food, shopping, canal walks, and evening energy are planned together.'},
                    @{After='kita-and-umeda-best-for-modern-osaka-and-transport'; File='osaka-umeda-architecture.png'; Alt='Modern Osaka Station and Umeda pedestrian architecture in daylight'; Caption='Umeda combines major transport connections with modern architecture, shopping, and skyline views.'}
                )
            }
            'osaka-2-day-itinerary' {
                @(
                    @{After='day-1-osaka-castle-nakanoshima-and-umeda'; File='osaka-2-day-castle.png'; Alt='Osaka Castle reflected in its moat during a calm morning'; Caption='Day one moves from Osaka Castle through the river districts before finishing in modern Umeda.'},
                    @{After='day-2-kuromon-market-namba-shinsekai-tennoji-shinsaibashi-and-dotonbori'; File='osaka-2-day-shinsekai.png'; Alt='Street-level evening view of Shinsekai and Tsutenkaku in Osaka'; Caption='Day two builds from markets and retro Shinsekai toward the brightest evening streets around Namba and Dotonbori.'}
                )
            }
            'osaka-3-day-itinerary' {
                @(
                    @{After='day-1-osaka-castle-nakanoshima-kitahama-and-umeda'; File='osaka-3-day-nakanoshima.png'; Alt='Nakanoshima riverside architecture in Osaka at golden hour'; Caption='The first day balances Osaka history with Nakanoshima, Kitahama, and Umeda city views.'},
                    @{After='day-3-schedule-summary'; File='osaka-3-day-minami.png'; Alt='Traveler walking beneath lanterns in an Osaka Minami food alley'; Caption='Three days create room to enjoy Minami at street level without compressing every southern district into one rushed evening.'}
                )
            }
            'where-to-stay-in-osaka' {
                @(
                    @{After='namba-best-overall-for-food-nightlife-and-a-first-osaka-experience'; File='where-to-stay-osaka-namba.png'; Alt='Boutique hotel entrance on a quieter Namba side street at dusk'; Caption='Namba places food and nightlife close by while quieter side streets can still provide a restful hotel base.'},
                    @{After='umeda-best-for-transport-day-trips-and-modern-osaka'; File='where-to-stay-osaka-umeda.png'; Alt='Traveler with suitcase near Osaka Station in Umeda'; Caption='Umeda is a practical base for regional trains, day trips, airport routes, and modern city conveniences.'}
                )
            }
            'kyoto-vs-osaka' {
                @(
                    @{After='when-kyoto-is-the-better-base'; File='kyoto-base-morning.png'; Alt='Quiet Kyoto machiya street and ryokan entrance in the early morning'; Caption='Kyoto is the stronger base when early temple access, traditional atmosphere, and calmer evenings matter most.'},
                    @{After='when-osaka-is-the-better-base'; File='osaka-base-night.png'; Alt='Warm lantern-lit food alley in Osaka at night'; Caption='Osaka is the stronger base for late dining, nightlife, and convenient regional transport connections.'}
                )
            }
            'iceland-travel-guide' {
                @(
                    @{After='understanding-iceland-s-main-regions'; File='iceland-travel-regions.png'; Alt='Iceland road trip landscape with mountains waterfalls and an open road'; Caption='A first Iceland route works best when regions are grouped realistically instead of treated as isolated attractions.'},
                    @{After='south-coast-one-of-the-strongest-first-time-regions'; File='iceland-south-coast.png'; Alt='Iceland South Coast waterfall black sand and coastal mountains'; Caption='The South Coast combines waterfalls, volcanic landscapes, ocean scenery, and useful overnight stops.'}
                )
            }
            'how-to-plan-a-trip-to-iceland' {
                @(
                    @{After='step-4-choose-your-priority-regions'; File='plan-iceland-route.png'; Alt='Travelers planning an Iceland road trip route with a map'; Caption='Choose priority regions before hotels so the route stays logical and adaptable.'},
                    @{After='step-10-check-official-conditions-during-planning'; File='iceland-weather-road-check.png'; Alt='Traveler checking Iceland weather and road conditions before driving'; Caption='Weather, road conditions, and safety information should remain separate checks throughout the trip.'}
                )
            }
            'how-many-days-do-you-need-in-iceland' {
                @(
                    @{After='is-5-days-enough-for-iceland'; File='iceland-five-day-route.png'; Alt='Five-day Iceland route through the Golden Circle and South Coast'; Caption='Five days can support a focused southwest and South Coast itinerary without forcing the full Ring Road.'},
                    @{After='what-does-a-10-day-ring-road-trip-feel-like'; File='iceland-ten-day-ring-road.png'; Alt='Ten-day Iceland Ring Road journey through varied landscapes'; Caption='Ten days give the Ring Road more room for weather, scenic stops, and balanced driving days.'}
                )
            }
            'best-time-to-visit-iceland' {
                @(
                    @{After='june-in-iceland'; File='iceland-summer-midnight-sun.png'; Alt='Iceland summer landscape under long golden midnight-sun light'; Caption='Long summer daylight supports broader road trips, hiking, and flexible sightseeing.'},
                    @{After='december-in-iceland'; File='iceland-winter-northern-lights.png'; Alt='Northern Lights above a snowy Iceland landscape in winter'; Caption='Winter offers dark skies and dramatic scenery but requires a smaller, more flexible route.'}
                )
            }
            'iceland-trip-cost' {
                @(
                    @{After='the-biggest-iceland-trip-expenses'; File='iceland-cost-breakdown.png'; Alt='Iceland travel budget planning with accommodation car fuel food and activities'; Caption='Accommodation, transport, food, and paid experiences create most of the Iceland trip budget.'},
                    @{After='how-to-spend-less-on-food-without-eating-poorly'; File='iceland-budget-road-trip.png'; Alt='Budget-conscious Iceland road trip with groceries and a compact rental car'; Caption='A focused route, practical vehicle, grocery stops, and selective activities can control costs without weakening the trip.'}
                )
            }
            'what-to-pack-for-iceland' {
                @(
                    @{After='the-most-important-iceland-packing-rule-use-layers'; File='iceland-layering-system.png'; Alt='Iceland clothing layers including base mid waterproof and insulating layers'; Caption='A flexible layering system handles wind, rain, cold, and changing activity levels better than one heavy outfit.'},
                    @{After='what-shoes-should-you-wear-in-iceland'; File='iceland-waterproof-footwear.png'; Alt='Waterproof hiking footwear on a wet Iceland trail near a waterfall'; Caption='Reliable footwear matters on wet paths, uneven ground, gravel, mud, and changing conditions.'}
                )
            }
            'iceland-travel-mistakes' {
                @(
                    @{After='7-leaving-no-buffer-for-icelandic-weather'; File='iceland-weather-mistake.png'; Alt='Iceland road traveler facing rapidly changing wind rain and visibility'; Caption='Weather buffers are part of the itinerary, not unused time.'},
                    @{After='31-treating-reynisfjara-like-a-normal-beach'; File='reynisfjara-ocean-safety.png'; Alt='Reynisfjara black sand beach viewed from a safe distance in Iceland'; Caption='Reynisfjara requires constant attention to the ocean, warning system, barriers, and local instructions.'}
                )
            }
            'iceland-3-day-itinerary' {
                @(
                    @{After='day-2-golden-circle'; File='iceland-3-day-golden-circle.png'; Alt='Adnan Ahmed visiting the Golden Circle during a three-day Iceland itinerary'; Caption='Day two uses the Golden Circle for a compact introduction to Icelandic history, geothermal activity, and waterfalls.'},
                    @{After='what-if-the-golden-circle-weather-is-bad'; File='iceland-3-day-south-coast.png'; Alt='Adnan Ahmed exploring Iceland South Coast waterfalls during a short itinerary'; Caption='The South Coast day should stay flexible enough for weather, road conditions, and safe time at each stop.'}
                )
            }
            'iceland-5-day-itinerary' {
                @(
                    @{After='day-2-golden-circle-and-continue-south'; File='iceland-5-day-golden-circle.png'; Alt='Adnan Ahmed on the Golden Circle before continuing into South Iceland'; Caption='Continuing south after the Golden Circle removes unnecessary backtracking and improves the five-day route.'},
                    @{After='day-4-v-k-to-j-kuls-rl-n-and-diamond-beach'; File='iceland-5-day-glacier-lagoon.png'; Alt='Adnan Ahmed at Jokulsarlon Glacier Lagoon during a five-day Iceland itinerary'; Caption='Day four reaches the glacier-lagoon region with enough time for the major southeast Iceland landscapes.'}
                )
            }
            'iceland-7-day-itinerary' {
                @(
                    @{After='day-4-v-k-to-skaftafell-j-kuls-rl-n-and-diamond-beach'; File='iceland-7-day-southeast.png'; Alt='Adnan Ahmed exploring southeast Iceland glacier scenery during a one-week trip'; Caption='A full week creates space for southeast Iceland without compressing every waterfall, glacier, and beach into one day.'},
                    @{After='day-5-southeast-iceland-to-borgarnes'; File='iceland-7-day-west.png'; Alt='Adnan Ahmed traveling through western Iceland toward Borgarnes'; Caption='The move west balances the route and positions the final days for western Iceland or Snaefellsnes.'}
                )
            }
            'iceland-10-day-itinerary' {
                @(
                    @{After='day-5-h-fn-through-the-eastfjords-to-egilssta-ir'; File='iceland-10-day-eastfjords.png'; Alt='Adnan Ahmed overlooking the Eastfjords on an Iceland Ring Road trip'; Caption='The Eastfjords deserve a measured driving day with time for coastlines, small towns, and changing weather.'},
                    @{After='day-6-egilssta-ir-to-lake-m-vatn'; File='iceland-10-day-myvatn.png'; Alt='Adnan Ahmed in the geothermal landscapes around Lake Myvatn'; Caption='North Iceland adds geothermal scenery and a very different character to the ten-day Ring Road route.'}
                )
            }
            'how-to-get-around-iceland' {
                @(
                    @{After='option-1-renting-a-car-in-iceland'; File='iceland-transport-rental-car.png'; Alt='Adnan Ahmed beside a rental car on a scenic Iceland road'; Caption='A rental car offers the greatest flexibility when the route, season, vehicle, and daily conditions are planned together.'},
                    @{After='option-3-organized-tours-in-iceland'; File='iceland-transport-guided-tour.png'; Alt='Adnan Ahmed joining a small guided tour in Iceland'; Caption='Guided tours are a practical alternative for travelers who want rural scenery without taking responsibility for driving.'}
                )
            }
            'driving-in-iceland' {
                @(
                    @{After='single-lane-bridges'; File='iceland-driving-single-lane-bridge.png'; Alt='Adnan Ahmed safely observing a single-lane bridge on an Iceland road'; Caption='Single-lane bridges require lower speed, clear communication, and patience with approaching traffic.'},
                    @{After='wind-is-one-of-iceland-s-biggest-driving-risks'; File='iceland-driving-wind-weather.png'; Alt='Adnan Ahmed checking difficult wind and weather beside an Iceland rental car'; Caption='Wind can change vehicle control, stopping decisions, and even the safe use of car doors.'}
                )
            }
            'renting-a-car-in-iceland' {
                @(
                    @{After='2wd-vs-4wd-comparison'; File='iceland-rental-2wd-vs-4wd.png'; Alt='Adnan Ahmed comparing a compact 2WD and four-wheel-drive vehicle in Iceland'; Caption='The correct choice between 2WD and 4WD depends on season, road type, route, space, and rental restrictions.'},
                    @{After='what-is-an-insurance-excess'; File='iceland-rental-insurance-check.png'; Alt='Adnan Ahmed inspecting an Iceland rental car and reviewing insurance'; Caption='Insurance decisions should account for excess, exclusions, gravel, wind, sand, tyres, and roadside assistance.'}
                )
            }
            'iceland-ring-road-guide' {
                @(
                    @{After='south-iceland-the-easiest-ring-road-section-to-understand'; File='iceland-ring-road-south.png'; Alt='Adnan Ahmed driving the South Iceland section of the Ring Road'; Caption='South Iceland has famous stops close to Route 1, but it still needs disciplined pacing and realistic overnight planning.'},
                    @{After='southeast-iceland-where-you-should-slow-down'; File='iceland-ring-road-southeast.png'; Alt='Adnan Ahmed overlooking glaciers and mountains in southeast Iceland'; Caption='Southeast Iceland rewards a slower day because glaciers, lagoons, coastline, and changing weather compete for time.'}
                )
            }
        }
        foreach ($placement in $placements) {
            $sectionPattern = '(<h2 class="section-title" id="' + [regex]::Escape($placement.After) + '">.*?</h2>)'
            $sectionImagePath = 'images/' + $placement.File
            $sectionImage = '<figure class="rp-post-image"><img src="' + $sectionImagePath + '" alt="' + (Html $placement.Alt) + '" width="1536" height="1024" loading="lazy" decoding="async"><figcaption>' + (Html $placement.Caption) + '</figcaption></figure>'
            $bodyHtml = ([regex]$sectionPattern).Replace($bodyHtml, ('$1' + $sectionImage), 1)
        }
        $bodyHtml = ([regex]'<details class="rp-faq-item">').Replace($bodyHtml, '<details class="rp-faq-item" open>', 1)
    }
    $toc=[Text.StringBuilder]::new(); $num=0; $seen=@{}; if($tokyoMode){$orderedToc=@($tocEntries | Sort-Object { $bodyHtml.IndexOf('id="' + $_.Id + '"') });foreach($entry in $orderedToc){$num++;[void]$toc.Append("<a href=`"#$($entry.Id)`"><span>$('{0:D2}' -f $num)</span><div>$(Html $entry.Text)</div></a>")}}else{foreach($h2 in $h2s){$num++;$base=Slug $h2.Text;$id=$base;$k=2;while($seen.ContainsKey($id)){$id="$base-$k";$k++};$seen[$id]=$true;[void]$toc.Append("<a href=`"#$id`"><span>$('{0:D2}' -f $num)</span><div>$(Html $h2.Text)</div></a>")}}
    $meta = ($post.Meta | ForEach-Object { "<span>$(Html $_)</span>" }) -join ''
    $html = @"
<!doctype html>
<!--
RoamPlans WordPress-ready post
Post title: $title
Original URL: https://roamplans.com/$($post.Slug)/
Category: $category
Author: Adnanahmed
The WordPress post content begins below.
-->
<!-- wp:html -->
<meta charset="utf-8"><meta content="width=device-width,initial-scale=1" name="viewport">
<title>$(Html $title) | RoamPlans</title>
<meta content="$(Html $desc)" name="description">
<style>$css</style>
<div class="progress" id="readingProgress"></div>
<div class="rp-final-grid">
<aside aria-label="RoamPlans discovery sidebar" class="rp-final-left"><!--ROAMPLANS_DISCOVERY_SIDEBAR--></aside>
<main class="rp-final-main"><article class="rp-post"><header class="rp-hero"><div class="hero-inner"><div class="eyebrow">RoamPlans &middot; $(Html $category)</div><h1>$(Html $title)</h1><div class="hero-deck">$(Html $desc)</div><div class="hero-meta">$meta</div></div></header><div class="article-card">$bodyHtml</div></article></main>
<aside aria-label="On this page" class="rp-final-right"><section class="side-card"><div class="side-kicker">On This Page</div><h2>$(Html $post.Short)</h2><nav class="toc">$($toc.ToString())</nav></section><section class="rp-note"><div class="side-kicker">RoamPlans Note</div><h3>$(Html $post.Note)</h3><p>$(Html $post.NoteText)</p></section></aside>
</div>
<div class="rp-final-author"><!--ROAMPLANS_AUTHOR_MODULE--></div>
<script>(function(){const progress=document.getElementById('readingProgress');const links=[...document.querySelectorAll('.toc a')];const sections=links.map(a=>document.querySelector(a.getAttribute('href'))).filter(Boolean);function update(){const max=document.documentElement.scrollHeight-innerHeight;progress.style.width=(max>0?(scrollY/max)*100:0)+'%';let current=sections[0];sections.forEach(section=>{if(section.getBoundingClientRect().top<150)current=section});links.forEach(link=>{link.classList.toggle('active',current&&link.getAttribute('href')==='#'+current.id)})}addEventListener('scroll',update,{passive:true});update()})();</script>
<!-- /wp:html -->
"@
    $out=Join-Path $OutputDir ($post.Slug+'.html')
    [IO.File]::WriteAllText($out,$html,[Text.UTF8Encoding]::new($false))
    Write-Host "Built $out"
}
