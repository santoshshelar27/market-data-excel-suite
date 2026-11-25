import { app, BrowserWindow, ipcMain, Menu, Tray } from 'electron';
import fs from 'node:fs';
import https from 'node:https';
import path from 'node:path';
import { WebSocketServer } from 'ws';

const PORT = Number(process.env.PORT || 3443);

function resolvePaths() {
  if (app.isPackaged) {
    return {
      distDir: path.join(process.resourcesPath, 'addin'),
      certDir: path.join(process.resourcesPath, 'certs')
    };
  }
  return {
    distDir: path.resolve(import.meta.dirname, '..', '..', 'excel-market-data-addin', 'dist', 'excel-market-data-addin', 'browser'),
    certDir: path.resolve(import.meta.dirname, '..', '..', 'excel-market-data-addin', 'electron-setup', 'certs')
  };
}

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.svg': 'image/svg+xml',
  '.xml': 'application/xml; charset=utf-8'
};

let mainWindow = null;
let tray = null;
let serverHandle = null;

function startAddinServer() {
  const { distDir, certDir } = resolvePaths();
  const keyPath = path.join(certDir, 'localhost.key');
  const certPath = path.join(certDir, 'localhost.crt');

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    throw new Error(`Missing TLS certs in ${certDir}. Run dev-certs.sh in excel-market-data-addin/electron-setup.`);
  }

  const server = https.createServer(
    {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath)
    },
    (req, res) => {
      const reqPath = (req.url || '/').split('?')[0];
      const safePath = path.normalize(reqPath === '/' ? '/index.html' : reqPath).replace(/^(\.\.[/\\])+/, '');
      let filePath = path.join(distDir, safePath);

      if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
        filePath = path.join(distDir, 'index.html');
      }

      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, {
        'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
        'Cache-Control': 'no-cache'
      });
      fs.createReadStream(filePath).pipe(res);
    }
  );

  const wss = new WebSocketServer({ server, path: '/ws' });
  const subscriptions = new Map(); // ws -> Set<string>

  const notifyStatus = () => {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    const allSymbols = new Set();
    for (const subs of subscriptions.values()) {
      for (const s of subs) allSymbols.add(s);
    }
    mainWindow.webContents.send('server-status', {
      port: PORT,
      clients: subscriptions.size,
      symbols: Array.from(allSymbols).sort()
    });
  };

  wss.on('connection', (ws) => {
    subscriptions.set(ws, new Set());
    notifyStatus();

    ws.on('message', (data) => {
      try {
        const msg = JSON.parse(data.toString());
        const symbol = typeof msg.symbol === 'string' ? msg.symbol.trim().toUpperCase() : '';
        if (!symbol) return;

        const clientSubs = subscriptions.get(ws);
        if (msg.type === 'subscribe') {
          clientSubs?.add(symbol);
        } else if (msg.type === 'unsubscribe') {
          clientSubs?.delete(symbol);
        }
        notifyStatus();
      } catch {
        // Ignore malformed JSON
      }
    });

    ws.on('close', () => {
      subscriptions.delete(ws);
      notifyStatus();
    });
  });

  function broadcastTick(tick) {
    const symbol = String(tick.symbol || '').toUpperCase();
    const payload = JSON.stringify(tick);
    for (const [client, subs] of subscriptions.entries()) {
      if (client.readyState === 1 && subs.has(symbol)) {
        client.send(payload);
      }
    }
  }

  // Simulated market data feed — replace with your live market data source
  const prices = new Map();
  const mockTimer = setInterval(() => {
    const allSymbols = new Set();
    for (const subs of subscriptions.values()) {
      for (const s of subs) allSymbols.add(s);
    }
    for (const symbol of allSymbols) {
      const prev = prices.get(symbol) ?? 100 + Math.random() * 100;
      const next = Number((prev + (Math.random() - 0.5) * 0.5).toFixed(2));
      prices.set(symbol, next);
      broadcastTick({
        symbol,
        price: next,
        bid: Number((next - 0.02).toFixed(2)),
        ask: Number((next + 0.02).toFixed(2)),
        volume: Math.floor(1000 + Math.random() * 50000)
      });
    }
  }, 250);

  server.listen(PORT, '127.0.0.1', () => {
    console.log(`MarketData HTTPS + WSS server listening on https://localhost:${PORT}`);
    notifyStatus();
  });

  return {
    broadcastTick,
    getStatus: () => {
      const allSymbols = new Set();
      for (const subs of subscriptions.values()) {
        for (const s of subs) allSymbols.add(s);
      }
      return {
        port: PORT,
        clients: subscriptions.size,
        symbols: Array.from(allSymbols).sort()
      };
    },
    stop: () => {
      clearInterval(mockTimer);
      wss.close();
      server.close();
    }
  };
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 480,
    height: 420,
    title: 'MarketData Server',
    webPreferences: {
      preload: path.join(import.meta.dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile(path.join(import.meta.dirname, 'index.html'));

  mainWindow.on('close', (event) => {
    if (!app.isQuitting) {
      event.preventDefault();
      mainWindow.hide();
    }
  });
}

function createTray() {
  const { distDir } = resolvePaths();
  const iconPath = path.join(distDir, 'assets', 'icon-16.png');
  if (!fs.existsSync(iconPath)) return;

  tray = new Tray(iconPath);
  const contextMenu = Menu.buildFromTemplate([
    { label: 'Show Status', click: () => mainWindow?.show() },
    {
      label: 'Quit',
      click: () => {
        app.isQuitting = true;
        app.quit();
      }
    }
  ]);
  tray.setToolTip(`MarketData Server (https://localhost:${PORT})`);
  tray.setContextMenu(contextMenu);
}

app.whenReady().then(() => {
  serverHandle = startAddinServer();
  createWindow();
  createTray();

  ipcMain.handle('get-server-status', () => serverHandle?.getStatus() ?? { port: PORT, clients: 0, symbols: [] });
});

app.on('before-quit', () => {
  app.isQuitting = true;
  serverHandle?.stop();
});
