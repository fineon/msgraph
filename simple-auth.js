/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
const express = require("express");
const msal = require('@azure/msal-node');
const graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
require('dotenv').config();

const session = require('express-session')

const SERVER_PORT = process.env.PORT || 3000;
const REDIRECT_URI = process.env.OAUTH_REDIRECT_URI;

// Create Express App and Routes
const app = express();

app.use(session({
    secret:'sec-key',
    resave: false,
    saveUninitialized: false,
    unset: 'destroy'
}));

// Before running the sample, you will need to replace the values in the config, 
// including the clientSecret
const config = {
    auth: {
        clientId: process.env.OAUTH_CLIENT_ID,
        authority: process.env.OAUTH_AUTHORITY,
        clientSecret: process.env.OAUTH_CLIENT_SECRET
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

// Create msal application object
const pca = new msal.ConfidentialClientApplication(config);


app.get('/', (req, res) => {
    const authCodeUrlParameters = {
      //TODO: look into one note scopes for every function init
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    // get url to sign user in and consent to scopes needed for application
    pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

function getAuthenticatedClient(msalClient, userId) {
  if (!msalClient || !userId) {
    throw new Error(
      `Invalid MSAL state. Client: ${msalClient ? 'present' : 'missing'}, User ID: ${userId ? 'present' : 'missing'}`);
  }

  // Initialize Graph client
  const client = graph.Client.init({
    // Implement an auth provider that gets a token
    // from the app's MSAL instance
    authProvider: async (done) => {
      try {
        // Get the user's account
        const account = await msalClient
          .getTokenCache()
          .getAccountByHomeId(userId);

        if (account) {
          // Attempt to get the token silently
          // This method uses the token cache and
          // refreshes expired tokens as needed
          const response = await msalClient.acquireTokenSilent({
            scopes: process.env.OAUTH_SCOPES.split(','),
            redirectUri: process.env.OAUTH_REDIRECT_URI,
            account: account
          });

          // First param to callback is the error,
          // Set to null in success case
          done(null, response.accessToken);
        }
      } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        done(err, null);
      }
    }
  });

  return client;
}

//working OK, got the auth token from the ?code param and saved in memory session
app.get('/auth/callback', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        //TODO: look into one note scopes for every function init
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        req.session.token = response;
        console.log(req.session.token);
        req.session.userId = response.account.homeAccountId;
        console.log(req.session.userId);
        res.sendStatus(200);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });

    //need to init msgrpah client with auth code or sth
    // graph.Client.api('/me/onenote/sections').get().then((info)=> console.log(info)).catch((err)=> console.error(err))
});

// having isues with 401 error, or code -1 unauthorized permission or missing auth token,. Prob a user permission thing on Azure AD or scope in .env or sth
app.get('/onenote',(req,res)=>{
  const apiTok = getAuthenticatedClient(pca, req.session.userId).api('/me/onenote/sections').get().catch((err)=> console.error(err));
  console.log(apiTok);
})

app.listen(SERVER_PORT, () => console.log(`Msal Node Auth Code Sample app listening on http://localhost:${SERVER_PORT}`))
