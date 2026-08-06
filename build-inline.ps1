# build-inline.ps1
# Rebuilds the self-contained index.html by inlining src/main.js and public/compass.html.
# Run this after editing any source file, then commit and push.

$root = $PSScriptRoot

$indexSrc    = Join-Path $root "index.html"
$mainJsSrc   = Join-Path $root "src\main.js"
$compassSrc  = Join-Path $root "public\compass.html"

Write-Host "Reading source files..."
$indexHtml   = [System.IO.File]::ReadAllText($indexSrc,   [System.Text.Encoding]::UTF8)
$mainJs      = [System.IO.File]::ReadAllText($mainJsSrc,  [System.Text.Encoding]::UTF8)
$compassHtml = [System.IO.File]::ReadAllText($compassSrc, [System.Text.Encoding]::UTF8)

# If index.html already has the JS inlined, strip back to the src= reference first
# so this script is idempotent (can be run multiple times safely).
if ($indexHtml -match '(?s)<script>\s*\n.*?var DIRECT=\[') {
    Write-Host "Stripping previously inlined script..."
    $indexHtml = $indexHtml -replace '(?s)<script>\s*\n.*?</script>(\s*</body>)', '<script src="src/main.js"></script>$1'
}
if ($indexHtml -match 'srcdoc=') {
    Write-Host "Stripping previously inlined srcdoc..."
    $indexHtml = $indexHtml -replace '(?s)<iframe srcdoc=".*?"(.*?)></iframe>', '<iframe src="public/compass.html"$1></iframe>'
}

# 1. Inline main.js
Write-Host "Inlining src/main.js..."
$inlineScript = "<script>`n" + $mainJs + "`n</script>"
$indexHtml = $indexHtml -replace '<script src="src/main\.js"></script>', $inlineScript

# 2. Inline compass.html as srcdoc
Write-Host "Inlining public/compass.html as srcdoc..."
$compassEscaped = $compassHtml -replace '&', '&amp;'
$compassEscaped = $compassEscaped -replace '"', '&quot;'
$srcdocIframe   = '<iframe srcdoc="' + $compassEscaped + '" style="width:100%;height:100%;border:none;" title="eLS Compass"></iframe>'
$indexHtml = $indexHtml -replace '<iframe src="public/compass\.html" style="width:100%;height:100%;border:none;" title="eLS Compass"></iframe>', $srcdocIframe

# Write output
[System.IO.File]::WriteAllText($indexSrc, $indexHtml, [System.Text.Encoding]::UTF8)
$size = [System.IO.File]::ReadAllBytes($indexSrc).Length
Write-Host "Done. index.html written ($size bytes)."
Write-Host ""
Write-Host "Next steps:"
Write-Host "  git add index.html"
Write-Host "  git commit -m 'Update portal content'"
Write-Host "  git push"
