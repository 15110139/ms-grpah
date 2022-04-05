require('dotenv').config()
import "isomorphic-fetch"
import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { MyAuthenticationProvider } from "./my-authentication-provider";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import USERS from "./users.json"
console.log("List user input from file:", USERS)
let clientOptions: ClientOptions = {
  authProvider: new MyAuthenticationProvider(),
};
const client = Client.initWithMiddleware(clientOptions);

async function addUsers(users: MicrosoftGraph.User[]) {
  try {
    const newUsers = await Promise.all(users.map(async (user: MicrosoftGraph.User) => await client.api('/users').post(user)))
    console.log("List user added:", newUsers)
  } catch (error) {
    console.error(error)
  }
}

async function listUser() {
  try {
    const users = await client.api('/users').get()
    console.log("List user:", users)
  } catch (error) {
    console.error(error)
  }
}

(async () => {
  await addUsers(USERS)
  await listUser()
})();



