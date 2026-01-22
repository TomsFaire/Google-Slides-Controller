# Package Companion Module for Testing
# Uses the official @companion-module/tools build command

Write-Host "Packaging Companion Module..." -ForegroundColor Green

# Go to companion module directory
Set-Location companion-module-gslide-opener

# Install dependencies if not present
if (-not (Test-Path "node_modules")) {
    Write-Host "Installing dependencies..." -ForegroundColor Yellow
    npm ci
} else {
    Write-Host "Dependencies already installed" -ForegroundColor Gray
}

# Run official build command
Write-Host "Building package with companion-module-build..." -ForegroundColor Yellow
npm run package

# Move the package to root and rename
Set-Location ..
$tgzFile = Get-ChildItem companion-module-gslide-opener\*.tgz -ErrorAction SilentlyContinue | Select-Object -First 1
if ($tgzFile) {
    Move-Item $tgzFile.FullName companion-module-gslide-opener.tgz -Force
}

if (Test-Path "companion-module-gslide-opener.tgz") {
    $size = (Get-Item "companion-module-gslide-opener.tgz").Length / 1MB
    $sizeRounded = [math]::Round($size, 2)
    Write-Host "Package created: companion-module-gslide-opener.tgz ($sizeRounded MB)" -ForegroundColor Green
    Write-Host ""
    Write-Host "To test in Companion:" -ForegroundColor Cyan
    Write-Host "1. Extract this .tgz file" -ForegroundColor White
    Write-Host "2. Point Developer modules path to the extracted 'pkg' folder" -ForegroundColor White
    Write-Host "3. Restart Companion" -ForegroundColor White
} else {
    Write-Host "Failed to create package" -ForegroundColor Red
}
