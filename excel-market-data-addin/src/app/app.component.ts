import { Component, inject } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { MarketWebSocketService } from './services/market-websocket.service';

@Component({
  selector: 'app-root',
  imports: [FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  readonly ws = inject(MarketWebSocketService);
  urlInput = this.ws.wsUrl();

  applyUrl(): void {
    this.ws.setUrl(this.urlInput);
  }
}
