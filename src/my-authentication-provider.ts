import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import axios, { AxiosRequestConfig } from 'axios';
import * as qs from 'qs'

const clientId = process.env.AUTH_CLIENT_ID;
const clientSecret = process.env.OAUTH_CLIENT_SECRET;
const tenantId = process.env.OAUTH_TENANT_ID;
const scope = process.env.OAUTH_SCOPES;


const data = qs.stringify({
    'grant_type': 'client_credentials',
    'client_id': clientId,
    'scope': scope,
    'client_secret': clientSecret
});
const config: AxiosRequestConfig = {
    method: 'get',
    url: `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
    },
    data: data
};

export class MyAuthenticationProvider implements AuthenticationProvider {
    public async getAccessToken(): Promise<string> {
        try {
            const data = await axios(config)
            return data.data.access_token
        } catch (error) {
            console.error(error)
            throw new Error(`Error when get token`)
        }
    }
}
