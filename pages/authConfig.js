 
// Config object to be passed to Msal on creation.
// For a full list of msal.js configuration parameters, 
// visit https://azuread.github.io/microsoft-authentication-library-for-js/docs/msal/modules/_authenticationparameters_.html
const msalConfig = {
  auth: {
    clientId: "e9535fba-33aa-41d5-b846-a4f88b7bd9d6",
    authority: "https://login.microsoftonline.com/common/",
    redirectUri: "https://6bdd120bed71.ngrok.io",
    navigateToLoginRequestUrl: false
  },
  cache: {
    cacheLocation: "localStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  }
};  
  
// Add here the scopes to request when obtaining an access token for MS Graph API
// for more, visit https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-core/docs/scopes.md
const loginRequest = {
  scopes: ["openid", "profile", "User.Read"],
  extraQueryParameters: {domain_hint: 'organizations'}
};

// Add here scopes for access token to be used at MS Graph API endpoints.
const tokenRequest = {
  scopes: ["Mail.Read", "Group.Read.All", "GroupMember.Read.All"],
  extraQueryParameters: {domain_hint: 'organizations'}
};