# Electron Market Data Server

Hosts the offline Angular 19 Excel Add-in over `https://localhost:3443` and streams JSON market data over `wss://localhost:3443/ws`.

## Development

1. Ensure the Angular add-in is built and dev certificates exist in `../excel-market-data-addin`:
   ```bash
   npm run build:addin
   ../excel-market-data-addin/electron-setup/dev-certs.sh
   ```
2. Install Electron dependencies and start:
   ```bash
   npm install
   npm start
   ```

## Build Single-Setup Installers

- **Windows (`.exe` NSIS):** `npm run dist:win`
- **macOS (`.pkg`):** `npm run dist:mac`
