/**
 * Handles webhook notifications from Microsoft Graph API
 */

import { http } from "@slflows/sdk/v1";
import type { AppOnHTTPRequestInput } from "@slflows/sdk/v1";
import { handleSubscriptionNotifications } from "./handleSubscriptionNotifications";

export async function handleWebhookNotification(
  input: AppOnHTTPRequestInput,
): Promise<void> {
  // Graph API sends validation requests with validationToken query parameter
  if (input.request.query?.validationToken) {
    await http.respond(input.request.requestId, {
      statusCode: 200,
      headers: {
        "Content-Type": "text/plain",
      },
      body: input.request.query.validationToken,
    });
    return;
  }

  const payload = input.request.body;

  if (payload?.value && Array.isArray(payload.value)) {
    try {
      await handleSubscriptionNotifications(payload.value, input.app.config);
    } catch (error: any) {
      console.error("Error handling notifications:", error);
    }

    await http.respond(input.request.requestId, {
      statusCode: 202,
      body: { status: "accepted" },
    });
  } else {
    await http.respond(input.request.requestId, {
      statusCode: 400,
      body: { error: "Invalid notification payload" },
    });
  }
}
