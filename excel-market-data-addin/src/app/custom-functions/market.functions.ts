import { getMarketWebSocketManager } from '../services/market-websocket.service';

/**
 * Streams the latest trade price for a given symbol.
 * @customfunction PRICE
 * @param symbol Ticker symbol (e.g., "AAPL").
 * @param invocation Custom function streaming invocation.
 */
export function price(
  symbol: string,
  invocation: CustomFunctions.StreamingInvocation<number | string>
): void {
  streamField(symbol, 'price', invocation);
}

/**
 * Streams the latest bid price for a given symbol.
 * @customfunction BID
 * @param symbol Ticker symbol (e.g., "AAPL").
 * @param invocation Custom function streaming invocation.
 */
export function bid(
  symbol: string,
  invocation: CustomFunctions.StreamingInvocation<number | string>
): void {
  streamField(symbol, 'bid', invocation);
}

/**
 * Streams the latest ask price for a given symbol.
 * @customfunction ASK
 * @param symbol Ticker symbol (e.g., "AAPL").
 * @param invocation Custom function streaming invocation.
 */
export function ask(
  symbol: string,
  invocation: CustomFunctions.StreamingInvocation<number | string>
): void {
  streamField(symbol, 'ask', invocation);
}

/**
 * Streams any named JSON field from the market data feed for a given symbol.
 * @customfunction FIELD
 * @param symbol Ticker symbol (e.g., "AAPL").
 * @param fieldName JSON property name (e.g., "volume", "changePct").
 * @param invocation Custom function streaming invocation.
 */
export function field(
  symbol: string,
  fieldName: string,
  invocation: CustomFunctions.StreamingInvocation<number | string | boolean>
): void {
  streamField(symbol, fieldName, invocation);
}

function streamField<T extends number | string | boolean>(
  symbol: string,
  fieldName: string,
  invocation: CustomFunctions.StreamingInvocation<T>
): void {
  if (!symbol || !symbol.trim()) {
    invocation.setResult('' as T);
    return;
  }

  const key = fieldName.trim();
  const manager = getMarketWebSocketManager();
  const unsubscribe = manager.subscribe(symbol, (tick) => {
    const value = tick[key];
    if (value !== undefined) {
      invocation.setResult(value as T);
    }
  });

  invocation.onCanceled = () => {
    unsubscribe();
  };
}

/**
 * Associates all streaming functions with the Excel CustomFunctions runtime.
 */
export function registerMarketCustomFunctions(): void {
  if (typeof CustomFunctions === 'undefined') {
    return;
  }
  CustomFunctions.associate('PRICE', price);
  CustomFunctions.associate('BID', bid);
  CustomFunctions.associate('ASK', ask);
  CustomFunctions.associate('FIELD', field);
}
