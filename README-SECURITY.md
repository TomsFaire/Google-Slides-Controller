# Security on macOS

## That security warning is normal

When you download and launch the app, macOS shows a security warning. This is expected. The app is safe—you can review the source code here on GitHub to verify it yourself.

The warning appears because the app uses self-signed code signing instead of Apple's official notarization service.

## First launch

**Right-click the app** and select **Open**. macOS will ask for confirmation—click **Open** again. That's the easiest way.

After the first launch, macOS remembers your choice and won't bother you again.

## App says it's "damaged"?

Sometimes macOS reports the app as "damaged" when extracted from a ZIP. This is just a quarantine flag. Fix it with:

```bash
xattr -cr "/path/to/Google Slides Opener.app"
```

Then try opening it again (right-click → Open).

## Alternative: System Settings

If right-clicking doesn't work, try:

1. Go to **System Settings** → **Privacy & Security**
2. Scroll down and look for "Google Slides Opener was blocked"
3. Click **Open Anyway**

## Why self-signed instead of notarized?

Apple's notarization costs $99/year, requires developer certificates, and involves submitting the app for review. For an open-source utility, self-signing is reasonable. Since the code is public, you can review it yourself to verify it's safe.

## See also (operators and admins)

- **[docs/PUBLIC-ACCESS.md](docs/PUBLIC-ACCESS.md)** — Quick Tunnel / shared Web UI links, **restricted vs full** Web UI (Settings hidden on the tunnel URL), and which API routes are blocked for remote viewers. Remote users are not shown an in-app notice about that restriction; configuration details are in that doc and the main [README.md](README.md).
- **Controller IP allowlist** and API exposure — covered from the main README and security-related sections there.
