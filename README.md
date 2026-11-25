# Market Data Excel Suite

An **offline, self-contained real-time market data solution** for Microsoft Excel (2019 / Microsoft 365+ on Windows and macOS), combining an **Angular 19 Office.js Excel Add-in** with an **Electron HTTPS + WebSocket (`wss://`) server** and **single-setup client installers** (no Microsoft Store required).

## Architecture

- **`excel-market-data-addin/`** — Angular 19 Excel Add-in using **Office.js Shared Runtime** (`@microsoft/office-js` bundled offline) and **Streaming Custom Functions** (`CustomFunctions.StreamingInvocation`) with 150ms tick throttling and reference-counted WebSocket subscriptions.
- **`electron-market-data-server/`** — Electron app that hosts the compiled Angular add-in over `https://localhost:3443`, runs the JSON WebSocket server at `wss://localhost:3443/ws`, and packages both into a single Windows (`.exe`) or macOS (`.pkg`) installer via `electron-builder`.

## Streaming Custom Functions

| Formula | Description |
| :--- | :--- |
| `=MARKET.PRICE("AAPL")` | Live trade price |
| `=MARKET.BID("AAPL")` | Live bid price |
| `=MARKET.ASK("AAPL")` | Live ask price |
| `=MARKET.FIELD("AAPL", "volume")` | Any JSON field from the tick payload |

## Quick Start

```bash
# 1. Install dependencies
npm --prefix excel-market-data-addin install
npm --prefix electron-market-data-server install
node electron-market-data-server/node_modules/electron/install.js

# 2. Generate local TLS certificates (Root CA + localhost)
./excel-market-data-addin/electron-setup/dev-certs.sh

# 3. Build the offline Angular add-in + functions.json
npm --prefix electron-market-data-server run build:addin

# 4. Start the Electron server
npm --prefix electron-market-data-server start
```

## Single-Setup Client Installers

From `electron-market-data-server/`:

- **Windows (`.exe` NSIS):** `npm run dist:win`
- **macOS (`.pkg`):** `npm run dist:mac`

For full certificate trust, sideloading, packaging, and troubleshooting instructions, see [SETUP_AND_INSTALLATION.md](./SETUP_AND_INSTALLATION.md).
