import { PublicClientApplication } from '@azure/msal-browser';

const msalConfig = {
  auth: {
    clientId: 'YOUR_CLIENT_ID_HERE', // Replace with your Azure AD application's client ID
    authority: 'https://login.microsoftonline.com/YOUR_TENANT_ID_HERE', // Replace with your Azure AD tenant ID
    redirectUri: 'http://localhost:3000/', // Ensure this matches the registered redirect URI
  },
};

const scopes = ['openid', 'profile', 'User.Read', 'Calendars.Read'];

let msalInstance;

export const initializeMsal = async () => {
  msalInstance = new PublicClientApplication(msalConfig);
  await msalInstance.handleRedirectPromise();
};

export const getAccessToken = async () => {
  try {
    if (!msalInstance) {
      await initializeMsal(); 
    }

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      throw new Error('No active account found');
    }

    const accessTokenPromises = accounts.map(async (account) => {
      const silentRequest = {
        scopes,
        account,
      };
      const response = await msalInstance.acquireTokenSilent(silentRequest);
      return response.accessToken;
    });

    const accessTokens = await Promise.all(accessTokenPromises);

    return accessTokens;
  } catch (error) {
    console.error('Failed to acquire access token:', error);
    throw error;
  }
};



export const refreshAccessToken = async (refreshToken, clientId) => {
  try {
    const response = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: new URLSearchParams({
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        client_id: clientId 
      })
    });

    const data = await response.json();
    const newAccessToken = data.access_token;

    return newAccessToken;
  } catch (error) {
    console.error('Failed to refresh access token:', error);
    throw error;
  }
};
