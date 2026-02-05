# Bundled Ghostscript

This app can run a bundled `gs` binary (Ghostscript) to convert PostScript (`.ps`) to PDF when the App Sandbox prevents executing Homebrew binaries.

## How to populate

Install Ghostscript:

```bash
brew install ghostscript
```

Then copy the `gs` binary into this folder **with the exact name `gs`**:

```bash
cp "$(brew --prefix ghostscript)/bin/gs" \
  "/Volumes/WDBlack4TB/Code/SendToOneNote/OneNoteHelperApp/Resources/ghostscript/gs"
chmod 755 "/Volumes/WDBlack4TB/Code/SendToOneNote/OneNoteHelperApp/Resources/ghostscript/gs"
```

Rebuild the app. At runtime, the app will extract the bundled `gs` to Application Support and execute it from there.
