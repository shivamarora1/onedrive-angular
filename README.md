 ## How to run this project:
 1. Transpile the source package `sdk-pnptimeline`
 ```
 cd sdk-pnptimeline
 npm install
 npm build
 ```
2. Create a new [AAD App Registration](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app?tabs=certificate), note the ID of the application.
3. Replace the clientId in `auth.ts` file.
 ```
 export const msalParams: Configuration = {
    auth: {
        authority: "https://login.microsoftonline.com/common",
        clientId: "<Your Client Id>",
        redirectUri: window.location.origin
    },
    cache: { cacheLocation: BrowserCacheLocation.LocalStorage }
}
```
4. Install the required dependencies
```
npm install --save
```
5. Run the application
```
npm run start
```
6. Building this project. Change script.start field in package.json as required.
```
npm run build
``` 

