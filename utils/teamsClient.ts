/**
 * Helper functions for Microsoft Teams Graph API calls
 */

import { kv } from "@slflows/sdk/v1";

export const TOKEN_KV_KEY = "ms_graph_access_token";
export const TOKEN_EXPIRY_KV_KEY = "ms_graph_token_expiry";

export interface GraphApiError {
  error: {
    code: string;
    message: string;
    innerError?: {
      code: string;
      message: string;
    };
  };
}

/**
 * Makes a call to the Microsoft Graph API
 * @param endpoint - The Graph API endpoint (e.g., "/teams/{id}/channels")
 * @param accessToken - The access token for authentication
 * @param options - Additional fetch options (method, body, etc.)
 * @returns The response data from the API
 */
export async function callGraphApi(
  endpoint: string,
  accessToken: string,
  options: {
    method?: string;
    body?: any;
    contentType?: string;
  } = {},
): Promise<any> {
  const { method = "GET", body, contentType = "application/json" } = options;

  const url = endpoint.startsWith("https://")
    ? endpoint
    : `https://graph.microsoft.com/v1.0${endpoint}`;

  const headers: Record<string, string> = {
    Authorization: `Bearer ${accessToken}`,
  };

  if (body && contentType) {
    headers["Content-Type"] = contentType;
  }

  const fetchOptions: RequestInit = {
    method,
    headers,
  };

  if (body) {
    fetchOptions.body = typeof body === "string" ? body : JSON.stringify(body);
  }

  const response = await fetch(url, fetchOptions);

  // Handle different response types
  const responseText = await response.text();
  let responseData;

  try {
    responseData = responseText ? JSON.parse(responseText) : {};
  } catch (e) {
    // If response is not JSON, return as text
    responseData = { text: responseText };
  }

  if (!response.ok) {
    const errorData = responseData as GraphApiError;
    const errorMessage =
      errorData.error?.message ||
      errorData.error?.code ||
      `Graph API request failed with status ${response.status}`;

    console.error(
      `Graph API Error [${method} ${endpoint}]:`,
      JSON.stringify(errorData, null, 2),
    );

    throw new Error(errorMessage);
  }

  return responseData;
}

/**
 * Gets a fresh access token from KV storage, refreshing if needed
 * @param config - The app config with Azure AD credentials
 * @returns A valid access token
 */
export async function getAccessToken(
  config: Record<string, any>,
): Promise<string> {
  // Try to get token from KV
  const tokenPair = await kv.app.get(TOKEN_KV_KEY);
  const expiryPair = await kv.app.get(TOKEN_EXPIRY_KV_KEY);

  if (tokenPair?.value && expiryPair?.value) {
    const tokenExpiry = parseInt(expiryPair.value);
    const bufferTime = 5 * 60 * 1000; // 5 minute buffer

    // Token is still valid
    if (Date.now() + bufferTime < tokenExpiry) {
      return tokenPair.value;
    }
  }

  // Token missing or expired - fetch new one
  const { accessToken, tokenExpiry } = await refreshAccessToken(config);

  // Store in KV
  await kv.app.set({ key: TOKEN_KV_KEY, value: accessToken });
  await kv.app.set({ key: TOKEN_EXPIRY_KV_KEY, value: tokenExpiry.toString() });

  return accessToken;
}

/**
 * Refreshes the access token by calling Azure AD
 * @param config - The app config with Azure AD credentials
 * @returns New access token and expiry time
 */
export async function refreshAccessToken(config: Record<string, any>): Promise<{
  accessToken: string;
  tokenExpiry: number;
}> {
  const { clientId, clientSecret, tenantId } = config;

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const tokenParams = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const tokenResponse = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: tokenParams.toString(),
  });

  if (!tokenResponse.ok) {
    const errorData = await tokenResponse.json();
    throw new Error(
      `Failed to refresh access token: ${errorData.error_description || errorData.error}`,
    );
  }

  const tokenData = await tokenResponse.json();
  const expiresIn = tokenData.expires_in || 3599;
  const tokenExpiry = Date.now() + expiresIn * 1000;

  return {
    accessToken: tokenData.access_token,
    tokenExpiry,
  };
}

/**
 * Creates a subscription to Graph API change notifications
 * @param resource - The resource to subscribe to (e.g., "/teams/{id}/channels/{id}/messages")
 * @param changeTypes - Array of change types to subscribe to (e.g., ["created", "updated"])
 * @param notificationUrl - The webhook URL to receive notifications
 * @param accessToken - The access token for authentication
 * @param expirationMinutes - How long the subscription should last (max 4230 for chat messages = ~3 days)
 * @returns The created subscription object
 */
export async function createSubscription(
  resource: string,
  changeTypes: string[],
  notificationUrl: string,
  lifecycleUrl: string,
  accessToken: string,
  expirationMinutes: number = 4230,
): Promise<any> {
  const expirationDateTime = new Date(
    Date.now() + expirationMinutes * 60 * 1000,
  ).toISOString();

  const subscriptionPayload = {
    changeType: changeTypes.join(","),
    notificationUrl: notificationUrl,
    lifecycleNotificationUrl: lifecycleUrl, // Separate endpoint for lifecycle events
    resource: resource,
    expirationDateTime: expirationDateTime,
    clientState: "flows-teams-app", // Used to verify notifications
  };

  return await callGraphApi("/subscriptions", accessToken, {
    method: "POST",
    body: subscriptionPayload,
  });
}

/**
 * Renews a subscription before it expires
 * @param subscriptionId - The ID of the subscription to renew
 * @param accessToken - The access token for authentication
 * @param expirationMinutes - How long to extend the subscription (max 4230)
 * @returns The updated subscription object
 */
export async function renewSubscription(
  subscriptionId: string,
  accessToken: string,
  expirationMinutes: number = 4230,
): Promise<any> {
  const expirationDateTime = new Date(
    Date.now() + expirationMinutes * 60 * 1000,
  ).toISOString();

  return await callGraphApi(`/subscriptions/${subscriptionId}`, accessToken, {
    method: "PATCH",
    body: {
      expirationDateTime: expirationDateTime,
    },
  });
}

/**
 * Deletes a subscription
 * @param subscriptionId - The ID of the subscription to delete
 * @param accessToken - The access token for authentication
 */
export async function deleteSubscription(
  subscriptionId: string,
  accessToken: string,
): Promise<void> {
  await callGraphApi(`/subscriptions/${subscriptionId}`, accessToken, {
    method: "DELETE",
  });
}

/**
 * Ensures a subscription exists for channel messages
 * Creates or renews the subscription and returns updated signal values
 * @param notificationUrl - The webhook URL to receive data notifications
 * @param lifecycleUrl - The webhook URL to receive lifecycle notifications
 * @param config - The app config with Azure AD credentials
 * @param signals - The app signals containing subscription state
 * @returns Updated subscription ID and expiry
 */
export async function ensureCentralSubscription(
  notificationUrl: string,
  lifecycleUrl: string,
  config: Record<string, any>,
  signals: {
    subscriptionId?: string;
    subscriptionExpiry?: string;
  },
): Promise<{
  subscriptionId: string;
  subscriptionExpiry: string;
  action: "created" | "renewed" | "valid";
}> {
  const accessToken = await getAccessToken(config);
  const bufferTime = 24 * 60 * 60 * 1000; // 24 hours buffer

  const subscriptionId = signals.subscriptionId;
  const subscriptionExpiry = signals.subscriptionExpiry
    ? parseInt(signals.subscriptionExpiry)
    : null;

  // Check if existing subscription is still valid
  if (subscriptionId && subscriptionExpiry) {
    const timeUntilExpiry = subscriptionExpiry - Date.now();

    // If subscription expires in more than 24 hours, it's still valid
    if (timeUntilExpiry > bufferTime) {
      console.log(
        "Subscription still valid for",
        Math.round(timeUntilExpiry / (60 * 60 * 1000)),
        "hours, no action needed",
      );
      return {
        subscriptionId,
        subscriptionExpiry: subscriptionExpiry.toString(),
        action: "valid",
      };
    }

    // Subscription expires within 24 hours - renew it
    try {
      console.log("Renewing existing subscription:", subscriptionId);
      const renewed = await renewSubscription(subscriptionId, accessToken);
      const newExpiry = new Date(renewed.expirationDateTime).getTime();
      console.log("Subscription renewed until:", renewed.expirationDateTime);
      return {
        subscriptionId,
        subscriptionExpiry: newExpiry.toString(),
        action: "renewed",
      };
    } catch (error: any) {
      console.warn(
        "Failed to renew subscription, will create new one:",
        error.message,
      );
      // Continue to create new subscription
    }
  }

  // Create new subscription for channel messages
  console.log("Creating new subscription for channel messages...");

  try {
    const subscription = await createSubscription(
      "/teams/getAllMessages",
      ["created", "updated"],
      notificationUrl,
      lifecycleUrl,
      accessToken,
      4230, // Max ~3 days
    );

    const expirationTime = new Date(subscription.expirationDateTime).getTime();
    console.log("Subscription created:", subscription.id);

    return {
      subscriptionId: subscription.id,
      subscriptionExpiry: expirationTime.toString(),
      action: "created",
    };
  } catch (error: any) {
    console.error("Failed to create subscription:", error.message);
    throw error;
  }
}
