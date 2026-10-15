/* global Office */

import {
  InteractionRequiredAuthError,
  createNestablePublicClientApplication,
} from "@azure/msal-browser";
import type { AuthenticationResult, IPublicClientApplication } from "@azure/msal-browser";

const CLIENT_ID = "b02153e0-a22d-4fef-8ed1-7fa0d3769a57";
const TENANT_ID = "5121092d-bf12-4971-adf1-9ae321302b3f";
const ACCESS_SCOPE = `api://${CLIENT_ID}/access_as_user`;

let msalInstancePromise: Promise<IPublicClientApplication> | undefined;

function getMsalInstance(): Promise<IPublicClientApplication> {
  if (!msalInstancePromise) {
    msalInstancePromise = createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      },
      cache: {
        cacheLocation: "sessionStorage",
      },
    });
  }

  return msalInstancePromise;
}

export function isNestedAppAuthSupported(): boolean {
  return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
}

export async function authenticateWithNaa(): Promise<AuthenticationResult> {
  if (!isNestedAppAuthSupported()) {
    throw new Error("NestedAppAuth 1.1 no está disponible en esta versión o cuenta de Outlook.");
  }

  const msalInstance = await getMsalInstance();
  const request = { scopes: [ACCESS_SCOPE] };

  try {
    return await msalInstance.ssoSilent(request);
  } catch (error: unknown) {
    if (error instanceof InteractionRequiredAuthError) {
      return msalInstance.acquireTokenPopup(request);
    }

    throw error;
  }
}

export function formatAuthenticationError(error: unknown): string {
  if (error instanceof Error) {
    const msalError = error as Error & { errorCode?: string; subError?: string };
    const details = [msalError.errorCode, msalError.subError].filter(Boolean).join(" / ");
    return `${error.name}${details ? ` (${details})` : ""}: ${error.message}`;
  }

  return String(error);
}
