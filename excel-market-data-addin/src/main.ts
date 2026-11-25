import { Office } from '@microsoft/office-js/excel';
import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { AppComponent } from './app/app.component';
import { registerMarketCustomFunctions } from './app/custom-functions/market.functions';
import { getMarketWebSocketManager } from './app/services/market-websocket.service';

function startApp(): void {
  registerMarketCustomFunctions();
  getMarketWebSocketManager().connect();

  bootstrapApplication(AppComponent, appConfig).catch((err) => {
    console.error('Failed to bootstrap Angular application:', err);
  });
}

if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
  Office.onReady(() => {
    startApp();
  });
} else {
  // Fallback when previewing in a standard browser outside Excel
  startApp();
}
