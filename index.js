const express = require("express");
const msal = require('@azure/msal-node');
const {SecretClient} = require('@azure/keyvault-secrets');
const {DefaultAzureCredential} = require('@azure/identity');

const SERVER_PORT = process.env.PORT || 3000;

const KEY_VAULT_URL = process.env['KEY_VAULT_URL'];
const SECRET_NAME = process.env['SECRET_NAME'];
const CLIENT_ID = process.env['CLIENT_ID'];
const TENANT_ID = process.env['AZURE_TENANT_ID'];


// Create Express App and Routes
const app = express();

app.listen(SERVER_PORT, () => console.log(`Msal Node Auth Code Sample app listening on port ${SERVER_PORT}!`))

const config = {
    auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + TENANT_ID,
        clientSecret: ""
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

new SecretClient(KEY_VAULT_URL, new DefaultAzureCredential()).getSecret(SECRET_NAME)
    .then((res) => {
        config.auth.clientSecret = res.value;
        // Create msal application object
        global.cca = new msal.ConfidentialClientApplication(config);
        return res;
    })
    .catch((error) => console.log(JSON.stringify(error)));


app.get('/', (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["user.read"],
        redirectUri: req.protocol + "://" + req.get("host") + "/redirect",
    };

    // get url to sign user in and consent to scopes needed for application
    cca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read"],
        redirectUri: "http://localhost:3000/redirect",
    };

    cca.acquireTokenByCode(tokenRequest).then((response) => {
        console.log("\nResponse: \n:", response);
        res.sendStatus(200);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});