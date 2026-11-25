# Installation & Single-Setup Distribution Guide (Windows & macOS)

This guide covers both:
1. **Developer Packaging:** Bundling the offline Angular 19 Excel Add-in and `manifest.xml` into your existing Electron app as a single installer (no Microsoft Store required).
2. **Client Installation:** What happens when the end user runs the single setup on Windows or macOS.

---

## Part 1: Prepare the Add-in & Electron Integration

### 1. Build the Offline Angular 19 Add-in
From the `excel-market-data-addin` root:

```bash
npm install
npm run build
```

This generates:
- `public/functions.json` (Custom Functions metadata)
- `dist/excel-market-data-addin/browser/` (Self-contained Angular app + offline `@microsoft/office-js` + `functions.json` + icons)

### 2. Generate TLS Certificates (One-Time or Per-Build)
Excel Desktop requires `https://localhost:3443` and `wss://localhost:3443/ws` to use a trusted certificate:

```bash
./electron-setup/dev-certs.sh
```

This outputs `rootCA.crt`, `localhost.crt`, and `localhost.key` into `electron-setup/certs/`.

### 3. Embed the Server in Your Electron Main Process
Copy `electron-setup/local-server.mjs` into your Electron project (and ensure `ws` is in your Electron `dependencies`), then call `startMarketAddinServer()` inside `app.whenReady()`:

```js
import path from 'node:path';
import { app } from 'electron';
import { startMarketAddinServer } from './local-server.mjs';

app.whenReady().then(() => {
  process.env.ADDIN_DIST_DIR = path.join(process.resourcesPath, 'addin');
  process.env.CERT_DIR = path.join(process.resourcesPath, 'certs');

  const addinServer = startMarketAddinServer();
  // Pipe your existing market data feed into addinServer.broadcastTick(tick)
});
```

### 4. Configure `electron-builder`
In your Electron project's `package.json`, bundle the Angular build, `manifest.xml`, and `certs` as `extraResources`, and wire up the platform installer scripts:

```json
{
  "build": {
    "appId": "com.yourcompany.marketdata",
    "productName": "MarketData",
    "extraResources": [
      {
        "from": "../excel-market-data-addin/dist/excel-market-data-addin/browser",
        "to": "addin"
      },
      {
        "from": "../excel-market-data-addin/manifest.xml",
        "to": "manifest.xml"
      },
      {
        "from": "../excel-market-data-addin/electron-setup/certs",
        "to": "certs"
      }
    ],
    "win": {
      "target": ["nsis"]
    },
    "nsis": {
      "oneClick": false,
      "perMachine": true,
      "allowElevation": true,
      "include": "../excel-market-data-addin/electron-setup/windows/installer.nsh"
    },
    "mac": {
      "target": ["pkg"]
    },
    "pkg": {
      "scripts": "../excel-market-data-addin/electron-setup/macos"
    }
  }
}
```

> **Important for macOS:** Update `APP_RESOURCES` inside `electron-setup/macos/postinstall` to match your `.app` bundle name (e.g., `/Applications/MarketData.app/Contents/Resources`).

---

## Part 2: Windows Installation Process

### Automated Actions Performed by the `.exe` (NSIS) Setup
When the client runs the installer (with Administrator privileges):
1. **Installs the Electron App** into `C:\Program Files\MarketData\`.
2. **Trusts the Local Root CA:** Runs `certutil -addstore -f "Root" "$INSTDIR\resources\certs\rootCA.crt"` so Edge WebView2 inside Excel trusts `https://localhost:3443` and `wss://localhost:3443/ws`.
3. **Creates the Local Add-in Catalog:** Copies `manifest.xml` to `C:\ProgramData\MarketDataAddin\excel-catalog\manifest.xml`.
4. **Registers the Trusted Catalog in Excel:** Writes the catalog path to:
   `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A4E8F120-9D34-4B1E-8C22-112233445566}`

### Client Steps on Windows
1. Double-click `MarketData-Setup.exe` and accept the UAC Administrator prompt.
2. Launch the **MarketData** Electron app (or leave it running in the system tray).
3. Open **Microsoft Excel** (2019 / Microsoft 365+):
   - Go to **Home → Add-ins** (or **Insert → My Add-ins → Shared Folder**).
   - Select **Real-Time Market Data** and click **Add**.
4. Test a streaming formula in any cell:
   ```excel
   =MARKET.PRICE("AAPL")
   ```

---

## Part 3: macOS Installation Process

### Automated Actions Performed by the `.pkg` Setup
When the client runs the `.pkg` installer (with Administrator password):
1. **Installs the Electron App** into `/Applications/MarketData.app`.
2. **Trusts the Local Root CA:** Runs `security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain` on `rootCA.crt` so WKWebView inside Excel trusts `https://localhost:3443` and `wss://localhost:3443/ws`.
3. **Sideloads `manifest.xml`:** Copies `manifest.xml` directly into the active user's Excel `wef` directory:
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml`

### Client Steps on macOS
1. Double-click `MarketData.pkg` and enter the macOS Administrator password when prompted.
2. Open **MarketData** from `/Applications` so the local server is running.
3. Open **Microsoft Excel** (2019 / Microsoft 365+):
   - The **Market Feed** button appears on the **Home** ribbon tab automatically (or under **Insert → My Add-ins**).
4. Test a streaming formula in any cell:
   ```excel
   =MARKET.PRICE("AAPL")
   ```

---

## Part 4: Troubleshooting

| Issue | Platform | Resolution |
| :--- | :--- | :--- |
| `#VALUE!` or `#BUSY!` stuck in cell | Both | Ensure the Electron app is running and listening on `https://localhost:3443`. Open the **Market Feed** task pane to check WebSocket status. |
| Certificate / blank task pane error | Windows | Verify `MarketData Local Root CA` is present in `certmgr.msc` → *Trusted Root Certification Authorities → Certificates*. |
| Certificate / blank task pane error | macOS | Open **Keychain Access** → **System** → verify `MarketData Local Root CA` is marked **Always Trust**. |
| Add-in not appearing in Excel | Windows | Verify `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\` has the catalog entry and `C:\ProgramData\MarketDataAddin\excel-catalog\manifest.xml` exists. Restart Excel completely. |
| Add-in not appearing in Excel | macOS | Verify `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml` exists. Restart Excel completely. |
| Stale `functions.json` or UI after update | Both | Clear the Office WebView cache (`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` on Windows, or `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` on macOS) and restart Excel. |
