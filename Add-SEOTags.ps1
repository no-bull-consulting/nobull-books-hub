# Add-SEOTags.ps1
# Batch-applies missing canonical, Open Graph, and twitter:image tags
# Run from project root:  cd C:\Users\edwar\code\nobull-website-source
#                         .\Add-SEOTags.ps1

$docsPath = Join-Path $PSScriptRoot "docs"

$pages = @(
    @{
        file      = "about.html"
        canonical = "https://nobull.consulting/about.html"
        ogTitle   = "About no~bull consulting - Built in the UK, for UK Sole Traders"
        ogDesc    = "no~bull consulting builds honest, fairly priced Google-native software for UK sole traders and micro-businesses. Founded by Edward Jenkins. GDPR by design, MTD-ready."
        noindex   = $false
    },
    @{
        file      = "app.html"
        canonical = "https://nobull.consulting/app.html"
        ogTitle   = "no~bull books - Client Portal"
        ogDesc    = "Sign in to your no~bull books account."
        noindex   = $true
    },
    @{
        file      = "founders.html"
        canonical = "https://nobull.consulting/founders.html"
        ogTitle   = "Founding Members - no~bull books"
        ogDesc    = "Be one of 20 UK businesses shaping the future of no~bull books. Free for 6 months, half-price forever."
        noindex   = $false
    },
    @{
        file      = "hmrc-mtd-statement.html"
        canonical = "https://nobull.consulting/hmrc-mtd-statement.html"
        ogTitle   = "HMRC Making Tax Digital (MTD) Recognised Software - no~bull books"
        ogDesc    = "no~bull books has applied for HMRC recognition for MTD for VAT and MTD for Income Tax Self Assessment (ITSA). Submit VAT returns and quarterly updates directly to HMRC."
        noindex   = $false
    },
    @{
        file      = "invitation.html"
        canonical = "https://nobull.consulting/invitation.html"
        ogTitle   = "Founding Member - no~bull books"
        ogDesc    = "Set up your no~bull books Founding Member account."
        noindex   = $true
    },
    @{
        file      = "privacy.html"
        canonical = "https://nobull.consulting/privacy.html"
        ogTitle   = "Privacy Policy - no~bull consulting"
        ogDesc    = "How no~bull consulting collects, uses, and protects your data. GDPR compliant by design."
        noindex   = $false
    },
    @{
        file      = "terms.html"
        canonical = "https://nobull.consulting/terms.html"
        ogTitle   = "Terms of Service - no~bull consulting"
        ogDesc    = "Terms and conditions governing use of no~bull consulting software and services."
        noindex   = $false
    },
    @{
        file      = "visa.html"
        canonical = "https://nobull.consulting/visa.html"
        ogTitle   = "no~bull visa"
        ogDesc    = "Secure sign-in for no~bull consulting."
        noindex   = $true
    }
)

$twitterImagePages = @(
    "contact.html",
    "index.html",
    "products.html",
    "services.html"
)

function Insert-BeforeCloseHead {
    param(
        [string]$filePath,
        [string]$insertion
    )
    $content = Get-Content $filePath -Raw -Encoding UTF8
    if ($content -notmatch '</head>') {
        Write-Warning "  No </head> found in $filePath - skipped."
        return
    }
    $updated = $content -replace '(?i)</head>', "$insertion`n</head>"
    [System.IO.File]::WriteAllText($filePath, $updated, [System.Text.Encoding]::UTF8)
}

Write-Host ""
Write-Host "=== Pass 1: canonical + OG tags ===" -ForegroundColor Cyan

foreach ($page in $pages) {
    $filePath = Join-Path $docsPath $page.file

    if (-not (Test-Path $filePath)) {
        Write-Warning "  $($page.file) not found - skipped."
        continue
    }

    $content = Get-Content $filePath -Raw -Encoding UTF8
    $changes = @()
    $lines = @()

    if ($page.noindex -and $content -notmatch 'name="robots"') {
        $lines += '  <meta name="robots" content="noindex, nofollow">'
        $changes += "robots noindex"
    }

    if ($content -notmatch 'rel="canonical"') {
        $lines += "  <link rel=`"canonical`" href=`"$($page.canonical)`">"
        $changes += "canonical"
    }

    $ogLines = @()
    if ($content -notmatch 'property="og:type"')        { $ogLines += '  <meta property="og:type"        content="website">' }
    if ($content -notmatch 'property="og:site_name"')   { $ogLines += '  <meta property="og:site_name"   content="no~bull consulting">' }
    if ($content -notmatch 'property="og:title"')       { $ogLines += "  <meta property=`"og:title`"       content=`"$($page.ogTitle)`">" }
    if ($content -notmatch 'property="og:description"') { $ogLines += "  <meta property=`"og:description`" content=`"$($page.ogDesc)`">" }
    if ($content -notmatch 'property="og:url"')         { $ogLines += "  <meta property=`"og:url`"         content=`"$($page.canonical)`">" }
    if ($content -notmatch 'property="og:image"')       { $ogLines += '  <meta property="og:image"       content="https://nobull.consulting/logo.png">' }
    if ($content -notmatch 'property="og:locale"')      { $ogLines += '  <meta property="og:locale"      content="en_GB">' }

    if ($ogLines.Count -gt 0) {
        $lines += ""
        $lines += "  <!-- OPEN GRAPH -->"
        $lines += $ogLines
        $changes += "OG tags"
    }

    if ($lines.Count -eq 0) {
        Write-Host "  $($page.file) - already up to date." -ForegroundColor DarkGray
        continue
    }

    Insert-BeforeCloseHead -filePath $filePath -insertion ($lines -join "`n")
    Write-Host "  $($page.file) - added: $($changes -join ', ')" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Pass 2: twitter:image ===" -ForegroundColor Cyan

$twitterImageTag = '  <meta name="twitter:image" content="https://nobull.consulting/logo.png">'

foreach ($filename in $twitterImagePages) {
    $filePath = Join-Path $docsPath $filename

    if (-not (Test-Path $filePath)) {
        Write-Warning "  $filename not found - skipped."
        continue
    }

    $content = Get-Content $filePath -Raw -Encoding UTF8

    if ($content -match 'twitter:image') {
        Write-Host "  $filename - twitter:image already present." -ForegroundColor DarkGray
        continue
    }

    Insert-BeforeCloseHead -filePath $filePath -insertion $twitterImageTag
    Write-Host "  $filename - added twitter:image" -ForegroundColor Green
}

Write-Host ""
Write-Host "All done. Commit with:" -ForegroundColor Cyan
Write-Host "  git add docs/" -ForegroundColor White
Write-Host "  git commit -m 'seo: add canonical, OG, and twitter:image tags'" -ForegroundColor White
Write-Host "  git push" -ForegroundColor White