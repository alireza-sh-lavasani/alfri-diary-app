// src/graphService.ts
import { Client } from "@microsoft/microsoft-graph-client";
import { IPublicClientApplication, AccountInfo } from "@azure/msal-browser";
import { loginRequest } from "./auth/authConfig";

export async function getProfilePhoto(
  instance: IPublicClientApplication,
  account: AccountInfo
): Promise<string | null> {
  const accessToken = await instance.acquireTokenSilent({
    ...loginRequest,
    account: account,
  });

  const graphClient = Client.init({
    authProvider: (done) => {
      done(null, accessToken.accessToken);
    },
  });

  try {
    const photo = await graphClient.api("/me/photo/$value").get();
    return URL.createObjectURL(photo);
  } catch (error) {
    console.error("Error fetching profile photo", error);
    return null;
  }
}
