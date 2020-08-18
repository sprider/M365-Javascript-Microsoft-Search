const msalConfig = {
  auth: {
    clientId: "",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:3000",
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  },
};

// Add scopes here for ID token to be used at Microsoft identity platform endpoints.
const loginRequest = {
  scopes: ["openid", "profile", "User.Read"],
};

// Add scopes here for access token to be used at Microsoft Graph API endpoints.
const tokenRequest = {
  scopes: [
    "Mail.Read",
    "Files.Read.All",
    "Calendars.Read",
    "ExternalItem.Read.All",
  ],
};
