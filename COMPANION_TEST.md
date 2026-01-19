# Testing Companion Module Locally

Before pushing to GitHub, you can test the companion module package locally.

## Quick Test

### 1. Package the module
```powershell
.\package-companion.ps1
```

This will:
- Install dependencies in the companion module
- Create `companion-module-gslide-opener.tgz` (same as GitHub Actions)

### 2. Verify the package
```powershell
.\test-companion-package.ps1
```

This will:
- Extract the .tgz
- Check that all required files are present
- Show you what's in the package
- Display version info

### 3. Test in Companion

**Option A: Extract and use dev path**
```powershell
# Extract the package
mkdir test-module
tar -xzf companion-module-gslide-opener.tgz -C test-module

# Point Companion's "Developer modules path" to: 
# C:\Users\nerif\Work Repos\gslide-opener\test-module
```

**Option B: Copy to Companion's user modules**
```powershell
# Extract directly to Companion's modules folder
# Usually: C:\Users\<username>\companion\module-local\<module-id>
```

## What to check

✅ **Config loads**: Can you see the Host and Port fields?
✅ **Config editable**: Can you change the Host and Port?
✅ **Actions available**: Do you see all the slide control actions?
✅ **Connection works**: Does it connect to the running app?

## Troubleshooting

**"No config data loaded"**
- Make sure `node_modules` is included in the package
- Run `.\test-companion-package.ps1` to verify

**"Module not found"**
- Check the extracted folder has all files
- Verify `companion/manifest.json` exists
- Restart Companion after adding/updating the module

**"Connection failed"**
- Make sure the Electron app is running
- Check the app is on `127.0.0.1:9595`
- Try changing Host/Port in Companion config
