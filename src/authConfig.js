export const msalConfig = {
  auth: {
    clientId: "<CLIENT_ID>",
    authority: "[login.microsoftonline.com](https://login.microsoftonline.com/)<TENANT_ID>",
    redirectUri: "[localhost](http://localhost:3000)"
  }
};

export const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All"]
};
