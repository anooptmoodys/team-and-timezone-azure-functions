import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { OnBehalfOfCredential } from '@azure/identity';

export const createDelegatedGraphClient = (accessToken: string): Client => {
    return Client.init({
        authProvider: (done) => {
            done(null, accessToken); // Use the user's access token
        },
    });
};

// delegated client using userImpersonationAccessToken
export const createDelegatedGraphClientWithOnBehalfOf = (userImpersonationAccessToken: string): Client => {
    const credential = new OnBehalfOfCredential({
        userAssertionToken: userImpersonationAccessToken,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        tenantId: process.env.TENANT_ID
    });

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ['https://graph.microsoft.com/.default'],
    });

    const graphClient = Client.initWithMiddleware({
        authProvider: authProvider,
    });

    return graphClient;
};