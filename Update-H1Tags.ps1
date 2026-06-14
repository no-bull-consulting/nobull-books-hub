# Update-H1Tags.ps1
# Updates H1 text on about, products, and services pages for SEO
# Run from project root:  .\Update-H1Tags.ps1

$docsPath = Join-Path $PSScriptRoot "docs"

function Replace-InFile {
    param(
        [string]$filePath,
        [string]$oldText,
        [string]$newText
    )
    $content = Get-Content $filePath -Raw -Encoding UTF8
    if ($content -notmatch [regex]::Escape($oldText)) {
        Write-Warning "  Target string not found in $filePath - skipped."
        return
    }
    $updated = $content.Replace($oldText, $newText)
    [System.IO.File]::WriteAllText($filePath, $updated, [System.Text.Encoding]::UTF8)
}

Write-Host ""
Write-Host "=== Updating H1 tags ===" -ForegroundColor Cyan

# --- about.html ---
$file = Join-Path $docsPath "about.html"
$old = "        <h1 style=`"color: var(--oxford-blue); font-family: var(--font-header);`">`n            honest software for the independent professional and micro-business.`n        </h1>"
$new = "        <h1 style=`"color: var(--oxford-blue); font-family: var(--font-header);`">`n            honest Google-native software for UK sole traders and micro-businesses.`n        </h1>"
Replace-InFile -filePath $file -oldText $old -newText $new
Write-Host "  about.html - H1 updated" -ForegroundColor Green

# --- products.html ---
$file = Join-Path $docsPath "products.html"
$old = "    <h1>Software that does the job.<br>Nothing more, nothing less.</h1>"
$new = "    <h1>no~bull books: UK accounting software for sole traders, freelancers and micro-businesses.</h1>`n    <p style=`"color: var(--slate); font-size: 18px; margin-top: 12px;`">Software that does the job. Nothing more, nothing less.</p>"
Replace-InFile -filePath $file -oldText $old -newText $new
Write-Host "  products.html - H1 updated (slogan kept as subtext)" -ForegroundColor Green

# --- services.html ---
$file = Join-Path $docsPath "services.html"
$old = "        <h1>Built on the platforms you already trust.</h1>"
$new = "        <h1>Bespoke Google Workspace development for UK businesses.</h1>`n        <p style=`"color: var(--slate); font-size: 18px; margin-top: 12px;`">Built on the platforms you already trust.</p>"
Replace-InFile -filePath $file -oldText $old -newText $new
Write-Host "  services.html - H1 updated (slogan kept as subtext)" -ForegroundColor Green

Write-Host ""
Write-Host "All done. Commit with:" -ForegroundColor Cyan
Write-Host "  git add docs/" -ForegroundColor White
Write-Host "  git commit -m 'seo: improve H1 tags on about, products and services'" -ForegroundColor White
Write-Host "  git push" -ForegroundColor White