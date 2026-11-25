# Complete Setup & Installation Guide

This guide covers everything needed to **build**, **package into a single client installer**, and **install** the offline Real-Time Market Data Excel Add-in + Electron WebSocket Server on **Windows** and **macOS** (without the Microsoft Store).

---

## 1. Project Structure

```text
market-data-excel-suite/
├── excel-market-data-addin/          # Angular 19 + Office.js + Custom Functions
│   ├── manifest.xml                  # Shared Runtime + Custom Functions manifest
│   ├── public/functions.json         # Auto-generated custom function metadata
│   └── electron-setup/
│       ├── dev-certs.sh              # Generates local Root CA + localhost TLS cert
│       ├── certs/                    # Generated rootCA.crt, localhost.crt, localhost.key
│       ├── windows/installer.nsh     # Windows NSIS post-install (cert + registry catalog)
│       └── macos/postinstall         # macOS .pkg post-install (cert + Excel wef folder)
└── electron-market-data-server/      # Electron app + HTTPS/WSS server + electron-builder
    ├── package.json                  # Build scripts (dist:win, dist:mac) & extraResources
    └── src/
        ├── main.mjs                  # Serves Angular dist over HTTPS + WSS on port 3443
        ├── preload.js
        └── index.html                # Server status UI
```

---

## 2. One-Time Developer Setup

### Step 1: Install Dependencies
From `market-data-excel-suite/`:

```bash
# 1. Angular Excel Add-in dependencies
npm --prefix excel-market-data-addin install

# 2. Electron Server dependencies
npm --prefix electron-market-data-server install

# 3. Download the Electron binary (if npm skipped postinstall scripts)
node electron-market-data-server/node_modules/electron/install.js
```

### Step 2: Generate Local TLS Certificates
Excel Desktop requires `https://localhost:3443` and `wss://localhost:3443/ws` to use a trusted certificate:

```bash
./excel-market-data-addin/electron-setup/dev-certs.sh
```

This creates `rootCA.crt`, `localhost.crt`, and `localhost.key` inside `excel-market-data-addin/electron-setup/certs/`.

### Step 3: Build the Offline Angular Add-in
```bash
npm --prefix electron-market-data-server run build:addin
```
This runs `custom-functions-metadata` to generate `functions.json` and compiles the Angular 19 app (including bundled `@microsoft/office-js`) into `excel-market-data-addin/dist/excel-market-data-addin/browser/`.

---

## 3. Running Locally in Development (Mac)

1. **Trust the dev Root CA (once):**
   ```bash
   sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain \
     excel-market-data-addin/electron-setup/certs/rootCA.crt
   ```
2. **Sideload `manifest.xml` into Excel (once):**
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   cp excel-market-data-addin/manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
   ```
3. **Start the Electron server:**
   ```bash
   npm --prefix electron-market-data-server start
   ```
4. **Open Microsoft Excel**, create a blank workbook, and test:
   ```excel
   =MARKET.PRICE("AAPL")
   =MARKET.BID("AAPL")
   =MARKET.ASK("AAPL")
   =MARKET.FIELD("AAPL", "volume")
   ```

---

## 4. Creating the Single-Setup Client Installers

`electron-builder` (configured in `electron-market-data-server/package.json`) bundles the Electron app, the compiled Angular add-in, `manifest.xml`, and `certs/` into **one installer** via `extraResources`.

### Build Windows Single Setup (`.exe`)
Run on **Windows** from `electron-market-data-server/`:
```bash
npm run dist:win
```
* **Output:** `electron-market-data-server/release/MarketData Setup 1.0.0.exe`
* **Included Hook:** `excel-market-data-addin/electron-setup/windows/installer.nsh`

### Build macOS Single Setup (`.pkg`)
Run on **macOS** from `electron-market-data-server/`:
```bash
npm run dist:mac
```
* **Output:** `electron-market-data-server/release/MarketData-1.0.0.pkg`
* **Included Hook:** `excel-market-data-addin/electron-setup/macos/postinstall`

> **Replacing the Mock Feed:** In `electron-market-data-server/src/main.mjs`, replace the `setInterval` mock generator with your live market data source calling `broadcastTick({ symbol, price, bid, ask, ... })` before building the production installers.

---

## 5. Client Installation Process (End-User)

### Windows (`.exe` NSIS Installer)
**What the installer does automatically (with Admin UAC):**
1. Installs **MarketData** into `C:\Program Files\MarketData\`.
2. Trusts `rootCA.crt` in the Windows Trusted Root Certification Authorities store (`certutil -addstore -f "Root"`).
3. Copies `manifest.xml` to `C:\ProgramData\MarketDataAddin\excel-catalog\manifest.xml`.
4. Registers that folder in `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{A4E8F120-9D34-4B1E-8C22-112233445566}`.

**Steps for the client:**
1. Run `MarketData Setup 1.0.0.exe` and accept the Administrator prompt.
2. Launch **MarketData** (it stays active in the system tray).
3. Open **Microsoft Excel (2019 / M365+)**.
4. Go to **Home → Add-ins** (or **Insert → My Add-ins → Shared Folder**), select **Real-Time Market Data**, and click **Add**.
5. Enter `=MARKET.PRICE("AAPL")` in any cell.

---

### macOS (`.pkg` Installer)
**What the installer does automatically (with Admin password):**
1. Installs **MarketData.app** into `/Applications/MarketData.app`.
2. Trusts `rootCA.crt` in `/Library/Keychains/System.keychain` (`security add-trusted-cert`).
3. Copies `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml`.

**Steps for the client:**
1. Double-click `MarketData-1.0.0.pkg` and enter the Administrator password.
2. Open **MarketData** from `/Applications`.
3. Open **Microsoft Excel (2019 / M365+)** — the **Market Feed** button appears on the **Home** tab.
4. Enter `=MARKET.PRICE("AAPL")` in any cell.

---

## 6. Troubleshooting

| Symptom | Platform | Fix |
| :--- | :--- | :--- |
| `#VALUE!` or `#BUSY!` in cell | Both | Confirm **MarketData** is running on port `3443`. Click **Market Feed** on Excel's Home tab to check WebSocket status. |
| Blank task pane / SSL error | Windows | Open `certmgr.msc` → **Trusted Root Certification Authorities → Certificates** and verify `MarketData Local Root CA` is listed. |
| Blank task pane / SSL error | macOS | Open **Keychain Access** → **System** and verify `MarketData Local Root CA` is set to **Always Trust**. |
| Add-in missing in Excel | Windows | Verify `C:\ProgramData\MarketDataAddin\excel-catalog\manifest.xml` exists and the `TrustedCatalogs` registry key is present. Restart Excel. |
| Add-in missing in Excel | macOS | Verify `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml` exists. Restart Excel. |
| Old UI or formulas cached | Both | Clear Excel's Office WebView cache (`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` on Windows, `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/` on macOS) and restart Excel. |
