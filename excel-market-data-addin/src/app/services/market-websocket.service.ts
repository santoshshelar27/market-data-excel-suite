import { Injectable, signal } from '@angular/core';

export type ConnectionState = 'disconnected' | 'connecting' | 'connected' | 'reconnecting';

export interface MarketTick {
  symbol: string;
  price?: number;
  bid?: number;
  ask?: number;
  [key: string]: string | number | boolean | undefined;
}

type TickListener = (tick: MarketTick) => void;

const DEFAULT_WS_URL = 'wss://localhost:3443/ws';
const STORAGE_KEY_WS_URL = 'market_addin_ws_url';
const FLUSH_INTERVAL_MS = 150;
const MAX_RECONNECT_DELAY_MS = 15000;

/**
 * Shared singleton WebSocket manager used by both the Angular Task Pane
 * and Excel Streaming Custom Functions inside the Office.js Shared Runtime.
 */
export class MarketWebSocketManager {
  readonly connectionState = signal<ConnectionState>('disconnected');
  readonly wsUrl = signal<string>(this.loadSavedUrl());
  readonly activeSymbols = signal<string[]>([]);
  readonly tickCount = signal<number>(0);
  readonly lastMessageAt = signal<string | null>(null);
  readonly lastError = signal<string | null>(null);

  private socket: WebSocket | null = null;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private flushTimer: ReturnType<typeof setInterval> | null = null;
  private reconnectAttempt = 0;
  private manualDisconnect = false;

  private readonly subscribers = new Map<string, Set<TickListener>>();
  private readonly latestTicks = new Map<string, MarketTick>();
  private readonly dirtySymbols = new Set<string>();

  constructor() {
    this.startFlushLoop();
  }

  private loadSavedUrl(): string {
    try {
      return localStorage.getItem(STORAGE_KEY_WS_URL) || DEFAULT_WS_URL;
    } catch {
      return DEFAULT_WS_URL;
    }
  }

  setUrl(url: string): void {
    const trimmed = url.trim();
    if (!trimmed) return;
    this.wsUrl.set(trimmed);
    try {
      localStorage.setItem(STORAGE_KEY_WS_URL, trimmed);
    } catch {
      // Ignore storage errors in restricted contexts
    }
    this.reconnect();
  }

  connect(): void {
    if (
      this.socket &&
      (this.socket.readyState === WebSocket.OPEN || this.socket.readyState === WebSocket.CONNECTING)
    ) {
      return;
    }

    this.manualDisconnect = false;
    this.clearReconnectTimer();
    this.connectionState.set(this.reconnectAttempt > 0 ? 'reconnecting' : 'connecting');
    this.lastError.set(null);

    try {
      const ws = new WebSocket(this.wsUrl());
      this.socket = ws;

      ws.onopen = () => {
        this.reconnectAttempt = 0;
        this.connectionState.set('connected');
        this.lastError.set(null);
        this.resubscribeAll();
      };

      ws.onmessage = (event: MessageEvent) => {
        this.handleRawMessage(event.data);
      };

      ws.onerror = () => {
        this.lastError.set(`WebSocket error connecting to ${this.wsUrl()}`);
      };

      ws.onclose = (event: CloseEvent) => {
        this.socket = null;
        if (this.manualDisconnect) {
          this.connectionState.set('disconnected');
          return;
        }
        if (!event.wasClean && !this.lastError()) {
          this.lastError.set(`Connection closed (code ${event.code})`);
        }
        this.scheduleReconnect();
      };
    } catch (err) {
      this.lastError.set(err instanceof Error ? err.message : String(err));
      this.scheduleReconnect();
    }
  }

  disconnect(): void {
    this.manualDisconnect = true;
    this.clearReconnectTimer();
    if (this.socket) {
      this.socket.close(1000, 'Client disconnected');
      this.socket = null;
    }
    this.connectionState.set('disconnected');
  }

  reconnect(): void {
    this.disconnect();
    this.reconnectAttempt = 0;
    this.connect();
  }

  /**
   * Subscribe an Excel custom function cell to a symbol.
   * Returns an unsubscribe callback for `invocation.onCanceled`.
   */
  subscribe(rawSymbol: string, listener: TickListener): () => void {
    const symbol = rawSymbol.trim().toUpperCase();
    if (!symbol) {
      return () => {};
    }

    let set = this.subscribers.get(symbol);
    const isFirstSubscriber = !set;
    if (!set) {
      set = new Set<TickListener>();
      this.subscribers.set(symbol, set);
      this.syncActiveSymbolsSignal();
    }
    set.add(listener);

    // Ensure WebSocket connects automatically when a formula is entered in Excel
    if (this.connectionState() === 'disconnected' && !this.manualDisconnect) {
      this.connect();
    }

    if (isFirstSubscriber) {
      this.sendJson({ type: 'subscribe', symbol });
    } else {
      const cached = this.latestTicks.get(symbol);
      if (cached) {
        listener(cached);
      }
    }

    return () => {
      const currentSet = this.subscribers.get(symbol);
      if (!currentSet) return;
      currentSet.delete(listener);
      if (currentSet.size === 0) {
        this.subscribers.delete(symbol);
        this.latestTicks.delete(symbol);
        this.dirtySymbols.delete(symbol);
        this.syncActiveSymbolsSignal();
        this.sendJson({ type: 'unsubscribe', symbol });
      }
    };
  }

  private handleRawMessage(raw: unknown): void {
    if (typeof raw !== 'string') return;
    try {
      const parsed = JSON.parse(raw) as Record<string, unknown>;
      const payload =
        parsed['data'] && typeof parsed['data'] === 'object'
          ? (parsed['data'] as Record<string, unknown>)
          : parsed;

      const rawSymbol = payload['symbol'] ?? parsed['symbol'];
      if (typeof rawSymbol !== 'string' || !rawSymbol.trim()) {
        return;
      }

      const symbol = rawSymbol.trim().toUpperCase();
      const existing = this.latestTicks.get(symbol) ?? { symbol };
      const updated: MarketTick = { ...existing, ...(payload as MarketTick), symbol };

      this.latestTicks.set(symbol, updated);
      this.dirtySymbols.add(symbol);
      this.tickCount.update((n) => n + 1);
      this.lastMessageAt.set(new Date().toLocaleTimeString());
    } catch {
      // Ignore non-JSON or malformed frames
    }
  }

  private startFlushLoop(): void {
    if (this.flushTimer) return;
    this.flushTimer = setInterval(() => {
      if (this.dirtySymbols.size === 0) return;
      const symbolsToFlush = Array.from(this.dirtySymbols);
      this.dirtySymbols.clear();

      for (const symbol of symbolsToFlush) {
        const tick = this.latestTicks.get(symbol);
        const listeners = this.subscribers.get(symbol);
        if (!tick || !listeners) continue;
        for (const listener of listeners) {
          listener(tick);
        }
      }
    }, FLUSH_INTERVAL_MS);
  }

  private resubscribeAll(): void {
    for (const symbol of this.subscribers.keys()) {
      this.sendJson({ type: 'subscribe', symbol });
    }
  }

  private sendJson(message: Record<string, unknown>): void {
    if (this.socket && this.socket.readyState === WebSocket.OPEN) {
      this.socket.send(JSON.stringify(message));
    }
  }

  private scheduleReconnect(): void {
    this.clearReconnectTimer();
    this.connectionState.set('reconnecting');
    const delay = Math.min(1000 * Math.pow(2, this.reconnectAttempt), MAX_RECONNECT_DELAY_MS);
    this.reconnectAttempt += 1;
    this.reconnectTimer = setTimeout(() => {
      this.connect();
    }, delay);
  }

  private clearReconnectTimer(): void {
    if (this.reconnectTimer) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
  }

  private syncActiveSymbolsSignal(): void {
    this.activeSymbols.set(Array.from(this.subscribers.keys()).sort());
  }
}

let sharedManager: MarketWebSocketManager | null = null;

export function getMarketWebSocketManager(): MarketWebSocketManager {
  if (!sharedManager) {
    sharedManager = new MarketWebSocketManager();
  }
  return sharedManager;
}

@Injectable({ providedIn: 'root' })
export class MarketWebSocketService {
  private readonly manager = getMarketWebSocketManager();

  readonly connectionState = this.manager.connectionState;
  readonly wsUrl = this.manager.wsUrl;
  readonly activeSymbols = this.manager.activeSymbols;
  readonly tickCount = this.manager.tickCount;
  readonly lastMessageAt = this.manager.lastMessageAt;
  readonly lastError = this.manager.lastError;

  connect(): void {
    this.manager.connect();
  }

  disconnect(): void {
    this.manager.disconnect();
  }

  reconnect(): void {
    this.manager.reconnect();
  }

  setUrl(url: string): void {
    this.manager.setUrl(url);
  }
}
