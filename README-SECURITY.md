# Security Notice for macOS Users

## Expected Behavior

**This security warning is expected and normal behavior** when downloading the app from GitHub. The app is safe to use - you can review the source code in this repository.

When you first launch the app after downloading, macOS will show a security warning because the app is not notarized by Apple. This is expected for apps that use ad-hoc code signing (self-signed) rather than Apple's notarization service.

## First Launch - Recommended Method

**Right-click the app** and select **"Open"** from the context menu. macOS will show a dialog asking if you want to open it - click **"Open"** in that dialog. This is the easiest way to launch the app for the first time.

## Alternative Method

If you double-click the app and see a security error:

1. Go to **System Settings** → **Privacy & Security**
2. Scroll down to find a message about "Google Slides Opener" being blocked
3. Click **"Open Anyway"** button

## After First Launch

Once you've opened the app the first time using either method above, macOS will remember your choice and you won't see the security warning again on subsequent launches.

## Why This Happens

This app uses ad-hoc code signing (self-signed) rather than Apple's notarization service. To use Apple's notarization:
- Requires an Apple Developer account ($99/year)
- Requires proper code signing certificates
- Requires submitting the app to Apple for review

For an open-source utility app, ad-hoc signing is a reasonable approach. The security warning is macOS's way of protecting users from potentially malicious software, but since this is an open-source app you can verify its safety by reviewing the code.
