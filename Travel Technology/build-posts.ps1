$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

function Encode([string]$text) { [System.Web.HttpUtility]::HtmlEncode($text) }
function Slug([string]$text) {
  $s = $text.ToLowerInvariant() -replace '[^a-z0-9]+','-'
  return $s.Trim('-')
}

function Get-ParagraphText($node, $ns) {
  (($node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
}

function Get-CellHtml($cell, $ns) {
  $parts = foreach ($p in $cell.SelectNodes('./w:p', $ns)) {
    $text = Get-ParagraphText $p $ns
    if ($text.Trim()) { Encode $text }
  }
  $parts -join '<br>'
}

function Get-DocModel([string]$path) {
  $zip = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    $entry = $zip.GetEntry('word/document.xml')
    $reader = [IO.StreamReader]::new($entry.Open())
    try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
  } finally { $zip.Dispose() }

  $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
  $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $items = [Collections.Generic.List[object]]::new()

  foreach ($node in $xml.SelectSingleNode('//w:body', $ns).ChildNodes) {
    if ($node.LocalName -eq 'p') {
      $text = Get-ParagraphText $node $ns
      if (-not $text.Trim()) { continue }
      $styleNode = $node.SelectSingleNode('./w:pPr/w:pStyle', $ns)
      $style = if ($styleNode) { $styleNode.GetAttribute('val','http://schemas.openxmlformats.org/wordprocessingml/2006/main') } else { 'Normal' }
      $num = $null -ne $node.SelectSingleNode('./w:pPr/w:numPr', $ns)
      $items.Add([pscustomobject]@{ Type='p'; Text=$text.Trim(); Style=$style; List=$num })
    } elseif ($node.LocalName -eq 'tbl') {
      $rows = [Collections.Generic.List[object]]::new()
      foreach ($tr in $node.SelectNodes('./w:tr', $ns)) {
        $cells = @($tr.SelectNodes('./w:tc', $ns) | ForEach-Object { Get-CellHtml $_ $ns })
        $rows.Add($cells)
      }
      $items.Add([pscustomobject]@{ Type='table'; Rows=$rows })
    }
  }
  return ,$items
}

$css = @'
:root{--ink:#1d2d2a;--text:#43514e;--muted:#6f7d79;--green:#1f6a5a;--dark:#155347;--mint:#eaf5f0;--sand:#f7f1e7;--gold:#d8a84e;--white:#fff;--line:#dfe9e4;--shadow:0 18px 45px rgba(25,62,53,.10)}
*{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:#fbfcfb;color:var(--text);font-family:Inter,Arial,Helvetica,sans-serif;line-height:1.78}a{color:var(--green);text-decoration:none}.progress{position:fixed;left:0;top:0;height:4px;background:linear-gradient(90deg,var(--gold),var(--green));width:0;z-index:9999}
.rp-final-grid{width:min(1480px,calc(100vw - 32px));margin:0 auto;display:grid;grid-template-columns:260px minmax(0,1fr) 270px;gap:28px;align-items:start}.rp-final-left{padding-top:24px}.rp-final-main{min-width:0}.rp-final-right{padding-top:24px;position:sticky;top:24px}.rp-post{overflow:hidden}.rp-hero{position:relative;background:linear-gradient(135deg,#123f37 0%,#1f6a5a 58%,#498777 100%);color:#fff;padding:72px 42px 80px;overflow:hidden}.rp-hero:before,.rp-hero:after{content:"";position:absolute;border-radius:50%;background:rgba(255,255,255,.07)}.rp-hero:before{width:360px;height:360px;right:-110px;top:-160px}.rp-hero:after{width:220px;height:220px;left:-90px;bottom:-110px}.hero-inner{max-width:900px;margin:auto;position:relative;z-index:2}.eyebrow{display:inline-flex;align-items:center;gap:9px;padding:8px 13px;border:1px solid rgba(255,255,255,.28);background:rgba(255,255,255,.09);border-radius:999px;font-size:13px;font-weight:800;letter-spacing:.08em;text-transform:uppercase}.rp-hero h1{max-width:850px;margin:24px 0 20px;font-family:Georgia,"Times New Roman",serif;font-size:clamp(46px,4.4vw,68px);line-height:1.03;letter-spacing:-.045em;color:#fff}.hero-deck{max-width:790px;margin:0;font-size:19px;line-height:1.65;color:rgba(255,255,255,.9)}.hero-meta{display:flex;flex-wrap:wrap;gap:12px 22px;margin-top:30px;font-size:14px;color:rgba(255,255,255,.82)}
.article-card{margin-top:-42px;position:relative;z-index:3;background:#fff;border:1px solid var(--line);border-radius:30px;padding:52px 58px;box-shadow:var(--shadow)}.article-card p{font-size:17px;margin:0 0 20px}.article-card>p:first-child{font-size:20px;color:#31433f}.section-title{scroll-margin-top:28px;font-family:Georgia,"Times New Roman",serif;font-size:clamp(30px,3.8vw,46px);line-height:1.13;letter-spacing:-.025em;color:var(--ink);margin:68px 0 24px;padding-top:4px}.section-title:before{content:"";display:block;width:48px;height:5px;border-radius:999px;background:linear-gradient(90deg,var(--gold),var(--green));margin-bottom:18px}.sub-title{scroll-margin-top:28px;font-size:23px;line-height:1.3;color:var(--ink);margin:38px 0 15px}.rp-list{list-style:none;padding:0;margin:22px 0 30px;display:grid;gap:11px}.rp-list li{display:flex;gap:13px;align-items:flex-start;padding:13px 15px;background:#f7faf8;border:1px solid var(--line);border-radius:14px}.rp-list li:before{content:'\2713';width:27px;height:27px;display:grid;place-items:center;flex:0 0 27px;border-radius:9px;background:var(--mint);color:var(--green);font-weight:900}.rp-list li p{margin:0;font-size:16px}.rp-callout{display:flex;gap:16px;margin:30px 0;padding:24px 25px;border-radius:20px;background:linear-gradient(135deg,#edf7f3,#f8fbfa);border:1px solid #cfe4db}.callout-icon{width:42px;height:42px;border-radius:14px;background:var(--green);color:#fff;display:grid;place-items:center;flex:0 0 42px;font-size:19px}.rp-callout p{margin:0;font-size:16px}.table-wrap{overflow-x:auto;margin:26px 0 34px;border:1px solid var(--line);border-radius:18px}table{width:100%;border-collapse:collapse;min-width:560px;background:#fff}th{background:var(--dark);color:#fff;text-align:left;padding:16px 18px;font-size:14px}td{padding:15px 18px;border-top:1px solid var(--line);font-size:15px;vertical-align:top}tr:nth-child(even) td{background:#f8faf9}
/* Premium Key Takeaways summary; scoped so regular article lists stay unchanged. */
#key-takeaways+.rp-list{counter-reset:takeaway;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin:24px 0 42px;padding:22px;background:linear-gradient(145deg,#edf7f3 0%,#f8f4e9 100%);border:1px solid #cfe4db;border-radius:24px;box-shadow:0 14px 34px rgba(25,62,53,.07)}
#key-takeaways+.rp-list li{counter-increment:takeaway;position:relative;display:grid;grid-template-columns:46px 1fr;gap:14px;align-items:start;min-height:92px;padding:18px;background:rgba(255,255,255,.92);border:1px solid rgba(31,106,90,.14);border-radius:17px;box-shadow:0 6px 16px rgba(25,62,53,.045)}
#key-takeaways+.rp-list li:before{content:counter(takeaway);width:42px;height:42px;display:grid;place-items:center;border-radius:13px;background:linear-gradient(145deg,var(--green),var(--dark));color:#fff;font-family:Georgia,"Times New Roman",serif;font-size:17px;font-weight:800;box-shadow:0 7px 15px rgba(31,106,90,.18)}
#key-takeaways+.rp-list li p{padding-top:6px;color:#344742;font-size:16px;line-height:1.55}
.side-card{background:#fff;border:1px solid var(--line);border-radius:22px;padding:24px;box-shadow:0 12px 30px rgba(25,62,53,.07)}.side-kicker{font-size:12px;font-weight:900;text-transform:uppercase;letter-spacing:.12em;color:var(--green);margin-bottom:8px}.side-card h2{margin:0 0 16px;font-family:Georgia,"Times New Roman",serif;font-size:27px;line-height:1.18;color:var(--ink)}.toc{display:grid;gap:7px;max-height:62vh;overflow:auto;padding-right:4px}.toc a{display:grid;grid-template-columns:31px 1fr;gap:9px;align-items:start;padding:9px 8px;border-radius:10px;color:#43514e;font-size:13px;line-height:1.35}.toc a:hover,.toc a.active{background:var(--mint);color:var(--dark)}.toc a span{font-size:11px;font-weight:900;color:var(--gold)}.rp-note{margin-top:20px;color:#fff;background:linear-gradient(145deg,var(--dark),var(--green));border-radius:22px;padding:24px;box-shadow:0 14px 34px rgba(25,62,53,.15)}.rp-note .side-kicker{color:#f5c96a}.rp-note h3{font-family:Georgia,"Times New Roman",serif;font-size:28px;line-height:1.15;margin:0 0 14px;color:#fff}.rp-note p{margin:0;color:rgba(255,255,255,.88);font-size:14px}
.rp-faq-list{display:grid;gap:16px;margin:30px 0 48px}.rp-faq-item{border:1px solid var(--line);border-radius:18px;background:#fff;overflow:hidden}.rp-faq-question{width:100%;display:flex;align-items:center;justify-content:space-between;gap:24px;padding:25px 27px;border:0;background:transparent;color:var(--ink);font:inherit;font-size:20px;font-weight:800;line-height:1.4;text-align:left;cursor:pointer}.rp-faq-icon{position:relative;width:22px;height:22px;flex:0 0 22px}.rp-faq-icon:before,.rp-faq-icon:after{content:"";position:absolute;left:50%;top:50%;width:16px;height:3px;border-radius:999px;background:var(--green);transform:translate(-50%,-50%);transition:transform .25s}.rp-faq-icon:after{transform:translate(-50%,-50%) rotate(90deg)}.rp-faq-item.is-open .rp-faq-icon:after{transform:translate(-50%,-50%)}.rp-faq-answer{display:grid;grid-template-rows:0fr;transition:grid-template-rows .3s}.rp-faq-item.is-open .rp-faq-answer{grid-template-rows:1fr}.rp-faq-answer-inner{overflow:hidden}.rp-faq-answer-inner p{margin:0;padding:0 27px 25px}.rp-final-author{width:min(1010px,calc(100vw - 32px));margin:56px auto 0}
/* RoamPlans branded premium details */
.rp-hero{isolation:isolate}.rp-hero .hero-inner:after{content:'RP';position:absolute;right:0;bottom:-38px;color:rgba(255,255,255,.055);font-family:Georgia,"Times New Roman",serif;font-size:150px;font-weight:900;line-height:1;letter-spacing:-.08em;z-index:-1}.hero-meta span{display:inline-flex;align-items:center;padding:7px 11px;border:1px solid rgba(255,255,255,.16);border-radius:999px;background:rgba(255,255,255,.07);backdrop-filter:blur(5px)}
.section-title:not(#key-takeaways){padding-bottom:15px;border-bottom:1px solid var(--line)}.sub-title{position:relative;padding-left:18px}.sub-title:before{content:'';position:absolute;left:0;top:.2em;width:5px;height:1.05em;border-radius:99px;background:linear-gradient(var(--gold),var(--green))}
.table-wrap{position:relative;box-shadow:0 12px 30px rgba(25,62,53,.06)}.table-wrap:before{content:'';display:block;height:4px;background:linear-gradient(90deg,var(--gold),var(--green))}.table-wrap th{position:sticky;top:0;z-index:1}.table-wrap tr{transition:background .18s ease}.table-wrap tbody tr:hover td,.table-wrap tr:hover td{background:#f0f7f4}
.rp-post-image{position:relative;margin:34px 0 42px;border:1px solid var(--line);border-radius:22px;overflow:hidden;background:linear-gradient(145deg,#f2f8f5,#faf6ed);box-shadow:0 14px 34px rgba(25,62,53,.08)}.rp-post-image img{display:block;width:100%;height:auto;aspect-ratio:16/10;object-fit:cover}.rp-post-image figcaption{padding:13px 17px;color:var(--muted);font-size:13px;background:#f8faf9;border-top:1px solid var(--line)}.rp-post-image.is-empty{min-height:360px;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:36px;text-align:center;border:2px dashed #b9d4c9;box-shadow:none}.rp-post-image.is-empty img{display:none}.rp-post-image.is-empty:before{content:'Image placeholder';display:grid;place-items:center;width:66px;height:66px;margin-bottom:16px;border-radius:20px;background:linear-gradient(145deg,var(--green),var(--dark));color:#fff;font-size:0;box-shadow:0 10px 22px rgba(31,106,90,.18)}.rp-post-image.is-empty:after{content:attr(data-image);max-width:90%;padding:9px 14px;border-radius:10px;background:#fff;color:var(--green);font-family:Consolas,monospace;font-size:13px;line-height:1.5;border:1px solid #d5e5de}.rp-post-image.is-empty figcaption{position:absolute;left:0;right:0;bottom:0}.rp-post-image.is-empty:before{background-image:linear-gradient(145deg,rgba(31,106,90,.96),rgba(21,83,71,.96));content:'+';font-size:34px;font-family:Arial,sans-serif;font-weight:300}
.rp-callout{position:relative;overflow:hidden;box-shadow:0 10px 25px rgba(25,62,53,.055)}.rp-callout:after{content:'';position:absolute;right:-32px;bottom:-46px;width:105px;height:105px;border-radius:50%;background:rgba(31,106,90,.055)}
.rp-faq-list{position:relative;padding:26px 22px 22px;background:linear-gradient(145deg,#f8fbf9 0%,#eef7f3 62%,#faf6ed 100%);border:1px solid #cfe4db;border-radius:26px;box-shadow:0 16px 36px rgba(25,62,53,.09);overflow:hidden}.rp-faq-list:before{content:'';position:absolute;left:0;right:0;top:0;height:5px;background:linear-gradient(90deg,var(--gold),var(--green) 55%,var(--dark))}.rp-faq-list:after{content:'?';position:absolute;right:-12px;top:-32px;color:rgba(31,106,90,.045);font-family:Georgia,"Times New Roman",serif;font-size:180px;font-weight:900;line-height:1;pointer-events:none}.rp-faq-item{position:relative;z-index:1;border-color:#d9e7e1;box-shadow:0 5px 14px rgba(25,62,53,.055);transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease}.rp-faq-item:hover{transform:translateY(-1px);border-color:#bdd8cd;box-shadow:0 9px 20px rgba(25,62,53,.09)}.rp-faq-item.is-open{border-color:#d9b55f;box-shadow:0 10px 24px rgba(25,62,53,.1)}.rp-faq-question{transition:color .2s ease,background .2s ease}.rp-faq-item.is-open .rp-faq-question{color:var(--dark);background:linear-gradient(90deg,#fbf7ed,#f1f8f5)}.rp-faq-answer-inner{background:#fff}.rp-faq-icon{border-radius:8px;background:var(--mint)}.rp-faq-item.is-open .rp-faq-icon{background:#f8edcf}
#final-thoughts{margin-bottom:0;padding:30px 32px 18px;border:1px solid #cfe4db;border-bottom:0;border-radius:24px 24px 0 0;background:linear-gradient(140deg,#edf7f3,#f8f4e9)}#final-thoughts:before{margin-bottom:14px}#final-thoughts+p{padding:0 32px 30px;margin:0 0 42px;border:1px solid #cfe4db;border-top:0;border-radius:0 0 24px 24px;background:linear-gradient(140deg,#edf7f3,#f8f4e9);color:#31433f;font-size:18px;box-shadow:0 14px 32px rgba(25,62,53,.07)}
.toc a{border-left:3px solid transparent;transition:background .18s ease,border-color .18s ease,transform .18s ease}.toc a:hover{transform:translateX(2px)}.toc a.active{border-left-color:var(--gold);font-weight:800}.back-to-top{position:fixed;right:24px;bottom:24px;width:48px;height:48px;display:grid;place-items:center;border:0;border-radius:16px;background:linear-gradient(145deg,var(--green),var(--dark));color:#fff;font-size:22px;box-shadow:0 12px 25px rgba(21,83,71,.25);cursor:pointer;opacity:0;visibility:hidden;transform:translateY(10px);transition:.22s ease;z-index:90}.back-to-top.is-visible{opacity:1;visibility:visible;transform:none}.back-to-top:hover{transform:translateY(-2px)}.back-to-top:focus-visible{outline:3px solid rgba(216,168,78,.5);outline-offset:3px}
@media(max-width:1200px){.rp-final-grid{grid-template-columns:230px minmax(0,1fr) 240px;gap:20px}.article-card{padding:42px 34px}}@media(max-width:1000px){.rp-final-grid{grid-template-columns:1fr;width:min(920px,calc(100vw - 24px))}.rp-final-left{order:2}.rp-final-main{order:1}.rp-final-right{order:3;position:static;padding-top:0}.toc{max-height:none}}@media(max-width:680px){.rp-final-grid{width:100%;padding:0 10px}.rp-hero{padding:54px 20px 68px}.rp-hero h1{font-size:40px}.rp-hero .hero-inner:after{font-size:100px;bottom:-22px}.hero-deck{font-size:17px}.article-card{padding:28px 20px;border-radius:22px}.section-title{font-size:34px}#key-takeaways+.rp-list{grid-template-columns:1fr;padding:14px;gap:11px;border-radius:19px}#key-takeaways+.rp-list li{grid-template-columns:40px 1fr;min-height:0;padding:15px}#key-takeaways+.rp-list li:before{width:38px;height:38px}.rp-faq-list{padding:13px;border-radius:20px}.rp-faq-question{padding:20px 18px;font-size:17px}.rp-faq-answer-inner p{padding:0 18px 20px}#final-thoughts{padding:24px 22px 14px}#final-thoughts+p{padding:0 22px 24px}.back-to-top{right:14px;bottom:14px;width:44px;height:44px;border-radius:14px}}
'@

$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
foreach ($file in Get-ChildItem $folder -Filter '*.docx' | Sort-Object Name) {
  $items = Get-DocModel $file.FullName
  $titleItem = $items | Where-Object { $_.Type -eq 'p' -and $_.Style -eq 'Heading1' } | Select-Object -First 1
  $title = $titleItem.Text
  $slug = Slug $title
  $intro = ($items | Where-Object { $_.Type -eq 'p' -and $_.Style -eq 'Normal' -and -not $_.List } | Select-Object -First 1).Text
  $description = if ($intro.Length -gt 155) { $intro.Substring(0,152).TrimEnd() + '...' } else { $intro }
  $toc = [Collections.Generic.List[string]]::new()
  $body = [Text.StringBuilder]::new()
  $inList = $false; $inFaq = $false; $faqOpen = $false; $seenTitle = $false; $headingIndex = 0; $tableImageAdded = $false; $sectionImageAdded = $false

  foreach ($item in $items) {
    if ($item.Type -eq 'table') {
      if ($inList) { [void]$body.AppendLine('</ul>'); $inList=$false }
      [void]$body.AppendLine('<div class="table-wrap"><table>')
      for ($r=0; $r -lt $item.Rows.Count; $r++) {
        $tag = if ($r -eq 0) { 'th' } else { 'td' }
        [void]$body.Append('<tr>')
        foreach ($cell in $item.Rows[$r]) { [void]$body.Append("<$tag>$cell</$tag>") }
        [void]$body.AppendLine('</tr>')
      }
      [void]$body.AppendLine('</table></div>')
      if (-not $tableImageAdded) {
        [void]$body.AppendLine('<figure class="rp-post-image" data-image="images/' + $slug + '-comparison.webp"><img src="images/' + $slug + '-comparison.webp" alt="' + (Encode ($title + ' comparison visual')) + '" loading="lazy" decoding="async" onerror="this.hidden=true;this.parentElement.classList.add(''is-empty'')"><figcaption>' + (Encode ($title + ' — comparison overview')) + '</figcaption></figure>')
        $tableImageAdded = $true
      }
      continue
    }

    if ($item.Style -eq 'Heading1' -and -not $seenTitle) { $seenTitle=$true; continue }
    if ($item.List) {
      if (-not $inList) { [void]$body.AppendLine('<ul class="rp-list">'); $inList=$true }
      [void]$body.AppendLine('<li><p>' + (Encode $item.Text) + '</p></li>')
      continue
    }
    if ($inList) { [void]$body.AppendLine('</ul>'); $inList=$false }

    # Some source documents mark Final Thoughts with the same heading level as
    # FAQ questions. Treat it as a new article section and close the accordion.
    if ($item.Text -eq 'Final Thoughts') {
      if ($faqOpen) { [void]$body.AppendLine('</div></div></div>'); $faqOpen=$false }
      if ($inFaq) { [void]$body.AppendLine('</div>'); $inFaq=$false }
      $id = 'final-thoughts'; $headingIndex++
      $toc.Add('<a href="#' + $id + '"><span>' + $headingIndex.ToString('00') + '</span>' + (Encode $item.Text) + '</a>')
      [void]$body.AppendLine('<h2 id="' + $id + '" class="section-title">' + (Encode $item.Text) + '</h2>')
      continue
    }

    if ($inFaq -and ($item.Style -eq 'Heading2' -or $item.Style -eq 'Heading3')) {
      if ($faqOpen) { [void]$body.AppendLine('</div></div></div>') }
      [void]$body.AppendLine('<div class="rp-faq-item"><button class="rp-faq-question" type="button" aria-expanded="false"><span>' + (Encode $item.Text) + '</span><span class="rp-faq-icon" aria-hidden="true"></span></button><div class="rp-faq-answer"><div class="rp-faq-answer-inner">')
      $faqOpen=$true
    } elseif ($item.Style -eq 'Heading2' -or $item.Style -eq 'Heading1') {
      if ($faqOpen) { [void]$body.AppendLine('</div></div></div>'); $faqOpen=$false }
      if ($inFaq) { [void]$body.AppendLine('</div>'); $inFaq=$false }
      $id = Slug $item.Text; $headingIndex++
      $toc.Add('<a href="#' + $id + '"><span>' + $headingIndex.ToString('00') + '</span>' + (Encode $item.Text) + '</a>')
      [void]$body.AppendLine('<h2 id="' + $id + '" class="section-title">' + (Encode $item.Text) + '</h2>')
      if ($headingIndex -eq 8 -and -not $sectionImageAdded -and $item.Text -ne 'Frequently Asked Questions') {
        [void]$body.AppendLine('<figure class="rp-post-image" data-image="images/' + $slug + '-section-guide.webp"><img src="images/' + $slug + '-section-guide.webp" alt="' + (Encode ($item.Text + ' visual guide')) + '" loading="lazy" decoding="async" onerror="this.hidden=true;this.parentElement.classList.add(''is-empty'')"><figcaption>' + (Encode ($title + ' - ' + $item.Text)) + '</figcaption></figure>')
        $sectionImageAdded = $true
      }
      if ($item.Text -eq 'Frequently Asked Questions') { [void]$body.AppendLine('<div class="rp-faq-list">'); $inFaq=$true }
    } elseif ($item.Style -eq 'Heading3') {
      if ($inFaq) {
        if ($faqOpen) { [void]$body.AppendLine('</div></div></div>') }
        [void]$body.AppendLine('<div class="rp-faq-item"><button class="rp-faq-question" type="button" aria-expanded="false"><span>' + (Encode $item.Text) + '</span><span class="rp-faq-icon" aria-hidden="true"></span></button><div class="rp-faq-answer"><div class="rp-faq-answer-inner">')
        $faqOpen=$true
      } else { [void]$body.AppendLine('<h3 class="sub-title">' + (Encode $item.Text) + '</h3>') }
    } elseif ($item.Text -match '^(App|Important|VPN|AI Planning|Translation|Gadget|Phone Preparation) Note:') {
      [void]$body.AppendLine('<div class="rp-callout"><div class="callout-icon" aria-hidden="true">!</div><div><p>' + (Encode $item.Text) + '</p></div></div>')
    } else { [void]$body.AppendLine('<p>' + (Encode $item.Text) + '</p>') }
  }
  if ($inList) { [void]$body.AppendLine('</ul>') }
  if ($faqOpen) { [void]$body.AppendLine('</div></div></div>') }
  if ($inFaq) { [void]$body.AppendLine('</div>') }

  $html = @"
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(Encode $title) | RoamPlans</title><meta name="description" content="$(Encode $description)"><style>$css</style></head><body>
<div class="progress" id="readingProgress"></div><div class="rp-final-grid">
<aside class="rp-final-left" aria-label="RoamPlans discovery sidebar"><!--ROAMPLANS_DISCOVERY_SIDEBAR--></aside>
<main class="rp-final-main"><article class="rp-post"><header class="rp-hero"><div class="hero-inner"><div class="eyebrow">RoamPlans &middot; Travel Technology</div><h1>$(Encode $title)</h1><div class="hero-meta"><span>&#10003; Practical guidance</span><span>&#8982; International travel</span><span>&#10022; Updated August 2026</span></div></div></header>
<div class="article-card">$body</div></article></main>
<aside class="rp-final-right" aria-label="Article navigation"><div class="side-card"><div class="side-kicker">In this guide</div><h2>On this page</h2><nav class="toc">$($toc -join "`n")</nav></div><div class="rp-note"><div class="side-kicker">RoamPlans tip</div><h3>Keep a backup</h3><p>Save essential travel information offline before you leave.</p></div></aside></div>
<div class="rp-final-author"><!--ROAMPLANS_AUTHOR_BOX--></div><button class="back-to-top" type="button" aria-label="Back to top" title="Back to top">&uarr;</button>
<script>const p=document.getElementById('readingProgress'),topButton=document.querySelector('.back-to-top');addEventListener('scroll',()=>{const d=document.documentElement;p.style.width=(d.scrollTop/(d.scrollHeight-d.clientHeight)*100)+'%';topButton.classList.toggle('is-visible',d.scrollTop>700)},{passive:true});topButton.addEventListener('click',()=>scrollTo({top:0,behavior:'smooth'}));document.querySelectorAll('.rp-faq-question').forEach(b=>b.addEventListener('click',()=>{const i=b.closest('.rp-faq-item'),o=i.classList.toggle('is-open');b.setAttribute('aria-expanded',o)}));const links=[...document.querySelectorAll('.toc a')],observer=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting)links.forEach(a=>a.classList.toggle('active',a.hash==='#'+e.target.id))}),{rootMargin:'-15% 0px -75%'});document.querySelectorAll('h2[id]').forEach(h=>observer.observe(h));</script></body></html>
"@
  $output = Join-Path $folder ($slug + '.html')
  [IO.File]::WriteAllText($output, $html, [Text.UTF8Encoding]::new($false))
  Write-Output "Created $([IO.Path]::GetFileName($output))"
}
