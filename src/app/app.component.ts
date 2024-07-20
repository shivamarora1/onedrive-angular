import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { BrowserAuthError } from '@azure/msal-browser';
import { getTokenWithScopes, getActiveAccount } from './auth';
import { CommonModule } from '@angular/common';
import { ButtonModule } from 'primeng/button';
import { PrimeNGConfig } from 'primeng/api';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule,ButtonModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent {
  title = 'onedrive-angular';
  userName: string = '';
  errorMsg: string = '';

  constructor(private primengConfig: PrimeNGConfig) {}

  ngOnInit() {
    this.primengConfig.ripple = true;
  }

  async onLogin() {
    try {
      this.errorMsg = '';
      await getTokenWithScopes([
        'user.read',
        'Sites.Read.All',
        'Files.Read.All',
        'offline_access',
      ]);
      const activeAcc = getActiveAccount();
      if (activeAcc) {
        this.userName = activeAcc?.username;
      } else {
        this.errorMsg = 'Unable to get active account info pls login again.';
      }
    } catch (e) {
      let errorMsg = 'something went wrong pls re login...';
      if (e instanceof BrowserAuthError) {
        errorMsg = e.errorMessage;
      }
      this.errorMsg = errorMsg;
      console.error(e);
    }
  }
}
