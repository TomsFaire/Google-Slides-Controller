# Package Companion Module for Testing
# This replicates what GitHub Actions does

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

# Go back to root
Set-Location ..

# Create the tarball (same as GitHub Actions)
Write-Host "Creating tarball..." -ForegroundColor Yellow
tar -czf companion-module-gslide-opener.tgz -C companion-module-gslide-opener .

if (Test-Path "companion-module-gslide-opener.tgz") {
    $size = (Get-Item "companion-module-gslide-opener.tgz").Length / 1MB
    Write-Host "✓ Package created: companion-module-gslide-opener.tgz ($([math]::Round($size, 2)) MB)" -ForegroundColor Green
    Write-Host ""
    Write-Host "To test in Companion:" -ForegroundColor Cyan
    Write-Host "1. Extract this .tgz file" -ForegroundColor White
    Write-Host "2. Place it in your Companion 'Developer modules path'" -ForegroundColor White
    Write-Host "3. Restart Companion" -ForegroundColor White
} else {
    Write-Host "✗ Failed to create package" -ForegroundColor Red
}
