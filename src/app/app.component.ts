import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { BrowserAuthError } from '@azure/msal-browser';
import { getTokenWithScopes, getActiveAccount, getToken } from './auth';
import { CommonModule } from '@angular/common';
import { ButtonModule } from 'primeng/button';
import { MessageService, PrimeNGConfig } from 'primeng/api';
import { OverlayPanelModule } from 'primeng/overlaypanel';
import { ToastModule } from 'primeng/toast';
import { TooltipModule } from 'primeng/tooltip';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { TableModule } from 'primeng/table';
import { formatBytes } from './utils/utils';
import {
  Client,
  AuthProvider,
  AuthProviderCallback,
  Options,
  GraphError,
} from '@microsoft/microsoft-graph-client';
import {
  Picker,
  Embed,
  IPickData,
  ResolveWithPicks,
  IFilePickerOptions,
  Popup,
  IPicker,
  LamdaAuthenticate,
  SPItem,
} from '@pnp/picker-api';

const requiredPermissions = [
  'user.read',
  'Sites.Read.All',
  'Files.Read.All',
  'offline_access',
];

const filePickerOptions: IFilePickerOptions = {
  sdk: '8.0',
  entry: {
    sharePoint: {},
  },
  authentication: {},
  messaging: {
    origin: window.location.origin,
    channelId: '27',
  },
  typesAndSources: {
    // filters: [".docx"],
    mode: 'files',
    pivots: {
      oneDrive: false,
      shared: true,
      sharedLibraries: true,
    },
  },
  selection: {
    mode: 'multiple',
    maxCount: 10,
  },
};

// const filePickerOptions: IFilePickerOptions = {
//   sdk: '8.0',
//   entry: {
//     oneDrive: {},
//   },
//   authentication: {},
//   messaging: {
//     origin: window.location.origin,
//     channelId: '27',
//   },
//   selection: {
//     mode: 'multiple',
//     maxCount: 5,
//   },
//   typesAndSources: {
//     mode: 'all',
//     pivots: {
//       recent: true,
//       oneDrive: true,
//     },
//   },
// };

@Component({
  selector: 'app-root',
  standalone: true,
  providers: [MessageService],
  imports: [
    ToastModule,
    TooltipModule,
    RouterOutlet,
    CommonModule,
    ButtonModule,
    OverlayPanelModule,
    TableModule,
  ],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent {
  title = 'onedrive-angular';
  userName: string = '';
  pickedItems!: SPItem[];

  constructor(
    private msgService: MessageService,
    private primengConfig: PrimeNGConfig
  ) { }

  ngOnInit() {
    this.primengConfig.ripple = true;
  }

  async onLogin() {
    try {
      await getTokenWithScopes(requiredPermissions);
      const activeAcc = getActiveAccount();
      if (activeAcc) {
        this.userName = activeAcc?.username;
        this.msgService.add({
          severity: 'success',
          detail: 'Successfully logged in as ' + this.userName,
        });
      } else {
        this.msgService.add({
          severity: 'danger',
          detail: 'Unable to get active account info pls login again.',
        });
      }
    } catch (e) {
      let errorMsg = 'something went wrong pls re login...';
      if (e instanceof BrowserAuthError) {
        errorMsg = e.errorMessage;
      }
      errorMsg = errorMsg;
      this.msgService.add({
        severity: 'danger',
        detail: errorMsg,
      });
      console.error(e);
    }
  }

  async onLogout() {
    console.log('Logout Clicked...');
  }

  onClose(): void {
    const ele = document.getElementsByTagName('iframe');
    if (ele && ele.length > 0) {
      ele[0].remove();
    }
  }

  formatBytes(b: number): string {
    return formatBytes(b);
  }

  async openSharepointOneDrive() {
    // * getting share point site URL before opening sharepoint file picker.
    try {
      this.pickedItems = null;
      const authProvider: AuthProvider = async (
        callback: AuthProviderCallback
      ) => {
        try {
          const accessToken = await getTokenWithScopes(requiredPermissions);
          callback(undefined, accessToken);
        } catch (e) {
          callback(e, '');
        }
      };
      let options: Options = {
        authProvider,
      };
      const client = Client.init(options);
      const result: MicrosoftGraph.Site = await client.api('/sites/root').get();
      const sharepointUrl = result.webUrl || '';

      if (sharepointUrl) {
        this.msgService.add({
          severity: 'success',
          detail: 'Your share point account is now connected',
        });
        await this.openFilePickerDialogBox(sharepointUrl);
      }
    } catch (e) {
      let errorMsg = 'Server error: Please try again.';
      if (
        e instanceof GraphError &&
        e.body.includes('Unable to find target address')
      ) {
        errorMsg = "Your account doesn't have access to OneDrive for Business";
      } else {
        console.error(e);
      }
      this.msgService.add({ severity: 'error', detail: errorMsg });
    }
  }

  async openFilePickerDialogBox(sharepointUrl: string = ''): Promise<void> {
    this.onClose();

    var iframe = document.createElement('iframe');
    iframe.src = '';
    iframe.width = '1000';
    iframe.height = '500';
    iframe.style.border = 'none';
    const iframeContainer: HTMLElement | null =
      document.getElementById('iframe-container');
    if (iframeContainer) {
      iframeContainer.appendChild(iframe);
    }

    let contentWindow = iframe.contentWindow;
    if (contentWindow) {
      let picker = Picker(contentWindow).using(
        ResolveWithPicks(),
        Popup(),
        LamdaAuthenticate(getToken)
      );

      picker.on.notification(function (this: IPicker, message) {
        // * uncomment for debugging
        // this.log("notification: " + JSON.stringify(message));
      });

      picker.on.log(function (this: IPicker, message, level) {
        // * uncomment for debugging
        // console.log(`log: [${level}] ${message}`);
      });

      const parentComp = this;
      picker.on.close(function (this: IPicker) {
        parentComp.onClose();
      });

      picker.on.error(function (this: IPicker, err) {
        this.log(`error: ${err}`);
      });

      (async () => {
        const baseUrl = sharepointUrl || 'https://onedrive.live.com';
        const results: IPickData | void = await picker.activate({
          baseUrl,
          options: filePickerOptions,
        });
        if (results) {
          this.pickedItems = results.items;
          console.log(this.pickedItems);
          parentComp.onClose();
        }
      })();
    }
  }
}
