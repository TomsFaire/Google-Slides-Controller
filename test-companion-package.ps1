# Test the companion module package
# Extracts and inspects the .tgz file

Write-Host "Testing Companion Package..." -ForegroundColor Green

if (-not (Test-Path "companion-module-gslide-opener.tgz")) {
    Write-Host "✗ Package not found. Run package-companion.ps1 first" -ForegroundColor Red
    exit 1
}

# Create temp directory
$tempDir = "test-package-temp"
if (Test-Path $tempDir) {
    Remove-Item -Recurse -Force $tempDir
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

Write-Host "Extracting package..." -ForegroundColor Yellow
tar -xzf companion-module-gslide-opener.tgz -C $tempDir

Write-Host ""
Write-Host "Package Contents:" -ForegroundColor Cyan
Get-ChildItem $tempDir -Recurse -Depth 2 | Select-Object FullName | ForEach-Object {
    $relativePath = $_.FullName.Replace("$PWD\$tempDir\", "")
    Write-Host "  $relativePath" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Checking critical files..." -ForegroundColor Yellow

$criticalFiles = @(
    "main.js",
    "actions.js",
    "package.json",
    "companion/manifest.json",
    "node_modules/@companion-module/base/package.json"
)

$allGood = $true
foreach ($file in $criticalFiles) {
    $fullPath = Join-Path $tempDir $file
    if (Test-Path $fullPath) {
        Write-Host "  ✓ $file" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $file (MISSING)" -ForegroundColor Red
        $allGood = $false
    }
}

Write-Host ""
if ($allGood) {
    Write-Host "✓ Package looks good!" -ForegroundColor Green
    
    # Show version info
    $manifest = Get-Content "$tempDir/companion/manifest.json" | ConvertFrom-Json
    Write-Host ""
    Write-Host "Module Info:" -ForegroundColor Cyan
    Write-Host "  Name: $($manifest.name)" -ForegroundColor White
    Write-Host "  Version: $($manifest.version)" -ForegroundColor White
    Write-Host "  ID: $($manifest.id)" -ForegroundColor White
} else {
    Write-Host "✗ Package has missing files!" -ForegroundColor Red
}

# Cleanup
Write-Host ""
Write-Host "Cleaning up..." -ForegroundColor Gray
Remove-Item -Recurse -Force $tempDir

Write-Host "Done!" -ForegroundColor Green
