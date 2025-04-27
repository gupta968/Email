const axios = require('axios');
const fs = require('fs');

// Microsoft Graph API configuration
const clientId = '<YOUR_CLIENT_ID>'; // Replace with your Azure AD app's client ID
const clientSecret = '<YOUR_CLIENT_SECRET>'; // Replace with your Azure AD app's client secret
const redirectUri = 'http://localhost'; // Redirect URI registered in your app
const tokenFilePath = './authTokens.json'; // File to store tokens

// Function to get authorization code from the user
async function getAuthorizationCode() {
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&response_mode=query&scope=offline_access%20https://graph.microsoft.com/Mail.Read.Shared`;

    console.log('Open this URL in your browser and log in:');
    console.log(authUrl);

    const rl = require('readline').createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    return new Promise((resolve) => {
        rl.question('Enter the authorization code: ', (code) => {
            rl.close();
            resolve(code);
        });
    });
}

// Function to get access and refresh tokens
async function getTokens(authCode) {
    const tokenUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('grant_type', 'authorization_code');
    params.append('code', authCode);
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('redirect_uri', redirectUri);

    const response = await axios.post(tokenUrl, params);
    return response.data;
}

// Function to refresh the access token using the refresh token
async function refreshAccessToken(refreshToken) {
    const tokenUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('grant_type', 'refresh_token');
    params.append('refresh_token', refreshToken);
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('redirect_uri', redirectUri);

    const response = await axios.post(tokenUrl, params);
    return response.data;
}

// Load tokens from file
function loadTokens() {
    if (fs.existsSync(tokenFilePath)) {
        return JSON.parse(fs.readFileSync(tokenFilePath, 'utf8'));
    }
    return null;
}

// Save tokens to file
function saveTokens(tokens) {
    fs.writeFileSync(tokenFilePath, JSON.stringify(tokens, null, 2));
}

// Main script
(async () => {
    try {
        let tokens = loadTokens();

        if (!tokens) {
            // Step 1: Get authorization code
            const authCode = await getAuthorizationCode();

            // Step 2: Get access and refresh tokens
            tokens = await getTokens(authCode);
            saveTokens(tokens);
        } else {
            // Step 3: Refresh access token if existing tokens are available
            console.log('Refreshing access token...');
            tokens = await refreshAccessToken(tokens.refresh_token);
            saveTokens(tokens);
        }

        console.log('Access Token:', tokens.access_token);

        // Use the access token to make API requests
        // Example: Fetch emails from shared mailbox
        const headers = {
            Authorization: `Bearer ${tokens.access_token}`,
        };
        const response = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages',
            { headers }
        );
        console.log('Emails:', response.data.value);
    } catch (error) {
        console.error('Error:', error.response ? error.response.data : error.message);
    }
})();
