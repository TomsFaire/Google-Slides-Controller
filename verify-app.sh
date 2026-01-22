#!/bin/bash
# Script to verify the app bundle structure before deployment

APP_PATH="$1"

if [ -z "$APP_PATH" ]; then
    echo "Usage: $0 <path-to-app>"
    echo "Example: $0 '/Applications/Google Slides Opener.app'"
    exit 1
fi

if [ ! -d "$APP_PATH" ]; then
    echo "❌ App not found at: $APP_PATH"
    exit 1
fi

echo "=== Verifying Google Slides Opener App Bundle ==="
echo ""

# Check main executable
if [ -f "$APP_PATH/Contents/MacOS/Google Slides Opener" ]; then
    echo "✅ Main executable found"
    file "$APP_PATH/Contents/MacOS/Google Slides Opener"
else
    echo "❌ Main executable missing!"
    exit 1
fi

echo ""

# Check Electron Framework
FRAMEWORK_PATH="$APP_PATH/Contents/Frameworks/Electron Framework.framework"
if [ -d "$FRAMEWORK_PATH" ]; then
    echo "✅ Electron Framework directory exists"
    
    # Check symlink
    if [ -L "$FRAMEWORK_PATH/Electron Framework" ]; then
        echo "✅ Electron Framework symlink exists"
        TARGET=$(readlink "$FRAMEWORK_PATH/Electron Framework")
        echo "   Points to: $TARGET"
        
        # Check if target exists
        if [ -f "$FRAMEWORK_PATH/$TARGET" ]; then
            echo "✅ Symlink target exists and is a file"
            file "$FRAMEWORK_PATH/$TARGET"
        else
            echo "❌ Symlink target does not exist or is not a file!"
            echo "   Expected: $FRAMEWORK_PATH/$TARGET"
            exit 1
        fi
    elif [ -f "$FRAMEWORK_PATH/Electron Framework" ]; then
        echo "✅ Electron Framework exists as direct file"
        file "$FRAMEWORK_PATH/Electron Framework"
    else
        echo "❌ Electron Framework not found!"
        exit 1
    fi
else
    echo "❌ Electron Framework directory missing!"
    exit 1
fi

echo ""

# Check Versions/Current symlink
if [ -L "$FRAMEWORK_PATH/Versions/Current" ]; then
    echo "✅ Versions/Current symlink exists"
    CURRENT_TARGET=$(readlink "$FRAMEWORK_PATH/Versions/Current")
    echo "   Points to: $CURRENT_TARGET"
    
    if [ -d "$FRAMEWORK_PATH/Versions/$CURRENT_TARGET" ]; then
        echo "✅ Current version directory exists"
    else
        echo "❌ Current version directory missing!"
        exit 1
    fi
else
    echo "❌ Versions/Current symlink missing!"
    exit 1
fi

echo ""
echo "=== All checks passed! App bundle structure is correct. ==="
