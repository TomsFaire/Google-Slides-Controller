# Security Notice for macOS Users

When you download the Google Slides Opener app from GitHub, macOS may show a security warning because the app is not notarized by Apple.

## Quick Solution (Easiest)

**Right-click the app** and select **"Open"** from the context menu. macOS will show a dialog asking if you want to open it - click **"Open"** in that dialog. This is easier than going to System Settings.

## Alternative Solution

If double-clicking shows a security error:

1. Go to **System Settings** → **Privacy & Security**
2. Scroll down to find a message about "Google Slides Opener" being blocked
3. Click **"Open Anyway"** button

## Why This Happens

This app uses ad-hoc code signing (self-signed) rather than Apple's notarization service. To use Apple's notarization:
- Requires an Apple Developer account ($99/year)
- Requires proper code signing certificates
- Requires submitting the app to Apple for review

For an open-source utility app, ad-hoc signing is a reasonable approach. The app is safe to use - you can review the source code in this repository.

## After First Launch

Once you've opened the app the first time, macOS will remember your choice and you won't see the security warning again.
