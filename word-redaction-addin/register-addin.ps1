# PowerShell script to register the Word Add-in for development
# Run this script as Administrator to sideload the add-in in Word Desktop

$manifestPath = Join-Path $PSScriptRoot "public\manifest.xml"
$regPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Word Add-in Registration Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if manifest exists
if (-not (Test-Path $manifestPath)) {
    Write-Host "ERROR: manifest.xml not found at: $manifestPath" -ForegroundColor Red
    Write-Host "Please make sure you're running this script from the project root directory." -ForegroundColor Red
    exit 1
}

Write-Host "Manifest found: $manifestPath" -ForegroundColor Green

# Create registry key if it doesn't exist
try {
    if (-not (Test-Path $regPath)) {
        Write-Host "Creating registry key..." -ForegroundColor Yellow
        New-Item -Path $regPath -Force | Out-Null
    }
    
    # Set the manifest path
    Write-Host "Registering add-in..." -ForegroundColor Yellow
    New-ItemProperty -Path $regPath -Name "UseManifest" -Value $manifestPath -PropertyType String -Force | Out-Null
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "SUCCESS! Add-in registered successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Make sure the dev server is running (npm run dev)" -ForegroundColor White
    Write-Host "2. Close all Word windows" -ForegroundColor White
    Write-Host "3. Open Word" -ForegroundColor White
    Write-Host "4. Look for 'Redaction' group on the Home ribbon" -ForegroundColor White
    Write-Host "5. Click 'Show Taskpane' to open the add-in" -ForegroundColor White
    Write-Host ""
    
} catch {
    Write-Host ""
    Write-Host "ERROR: Failed to register add-in" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Make sure you're running PowerShell as Administrator." -ForegroundColor Yellow
    exit 1
}
