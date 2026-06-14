# Add-PersonSchema.ps1
# Adds Person + Organization JSON-LD structured data to about.html
# Run from project root:  .\Add-PersonSchema.ps1

$filePath = Join-Path $PSScriptRoot "docs\about.html"

if (-not (Test-Path $filePath)) {
    Write-Error "about.html not found at $filePath"
    exit 1
}

$content = Get-Content $filePath -Raw -Encoding UTF8

if ($content -match 'application/ld\+json') {
    Write-Host "Structured data already present in about.html - skipped." -ForegroundColor DarkGray
    exit 0
}

$schema = @'

  <!-- STRUCTURED DATA - Person + Organization -->
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@graph": [
      {
        "@type": "Person",
        "@id": "https://nobull.consulting/#edward-jenkins",
        "name": "Edward Jenkins",
        "jobTitle": "Founder",
        "worksFor": {
          "@id": "https://nobull.consulting/#organization"
        },
        "url": "https://nobull.consulting/about.html",
        "sameAs": [
          "https://www.linkedin.com/in/edward-jenkins-aa4ab42a/"
        ]
      },
      {
        "@type": "Organization",
        "@id": "https://nobull.consulting/#organization",
        "name": "no~bull consulting",
        "url": "https://nobull.consulting",
        "logo": "https://nobull.consulting/logo.png",
        "foundingDate": "2024",
        "founder": {
          "@id": "https://nobull.consulting/#edward-jenkins"
        },
        "areaServed": {
          "@type": "Country",
          "name": "United Kingdom"
        },
        "description": "no~bull consulting builds honest, fairly priced Google-native software for UK sole traders and micro-businesses. GDPR by design, MTD-ready.",
        "contactPoint": {
          "@type": "ContactPoint",
          "email": "hello@nobull.consulting",
          "contactType": "customer support"
        }
      }
    ]
  }
  </script>
'@

$updated = $content -replace '(?i)</head>', "$schema`n</head>"
[System.IO.File]::WriteAllText($filePath, $updated, [System.Text.Encoding]::UTF8)

Write-Host ""
Write-Host "=== Person schema added to about.html ===" -ForegroundColor Cyan
Write-Host "  Person: Edward Jenkins, Founder" -ForegroundColor Green
Write-Host "  Organization: no~bull consulting" -ForegroundColor Green
Write-Host ""
Write-Host "NOTE: update the LinkedIn URL in the schema if needed:" -ForegroundColor Yellow
Write-Host "  https://www.linkedin.com/in/edwardjenkins" -ForegroundColor Yellow
Write-Host ""
Write-Host "Commit with:" -ForegroundColor Cyan
Write-Host "  git add docs/about.html" -ForegroundColor White
Write-Host "  git commit -m 'seo: add Person and Organization schema to about.html'" -ForegroundColor White
Write-Host "  git push" -ForegroundColor White