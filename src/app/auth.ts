import { PublicClientApplication, Configuration, SilentRequest, BrowserCacheLocation, AccountInfo } from "@azure/msal-browser";
import { combine } from "@pnp/core";
import { IAuthenticateCommand } from "@pnp/picker-api/dist";

export const msalParams: Configuration = {
    auth: {
        authority: "https://login.microsoftonline.com/common",
        clientId: "3242284e-c36d-4ac7-8bde-d02811524161",
        redirectUri: window.location.origin
    },
    cache: { cacheLocation: BrowserCacheLocation.LocalStorage }
}

const app = new PublicClientApplication(msalParams);

export async function getToken(command: IAuthenticateCommand): Promise<string> {
    return getTokenWithScopes([`${combine(command.resource, ".default")}`]);
}

export async function getTokenWithScopes(scopes: string[], additionalAuthParams?: Omit<SilentRequest, "scopes">): Promise<string> {
    await app.initialize()

    let accessToken = "";
    const authParams = { scopes, ...additionalAuthParams };

    try {
        const resp = await app.acquireTokenSilent(authParams!);
        accessToken = resp.accessToken;
    } catch (e) {
        const resp = await app.loginPopup(authParams!);
        app.setActiveAccount(resp.account);

        if (resp.idToken) {
            const resp2 = await app.acquireTokenSilent(authParams!);
            accessToken = resp2.accessToken;

        } else {
            throw e;
        }
    }
    return accessToken;
}

export function getActiveAccount(): AccountInfo | null {
    return app.getActiveAccount()

}