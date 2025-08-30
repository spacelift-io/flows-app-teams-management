/**
 * Handles lifecycle notifications from Microsoft Graph API
 */

import { http, lifecycle } from "@slflows/sdk/v1";
import type { AppOnHTTPRequestInput } from "@slflows/sdk/v1";

export async function handleLifecycleNotification(
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
    const firstNotification = payload.value[0];
    if (firstNotification?.lifecycleEvent) {
      if (firstNotification.lifecycleEvent === "reauthorizationRequired") {
        console.warn("Subscription requires reauthorization");
        await lifecycle.sync();
      } else if (firstNotification.lifecycleEvent === "subscriptionRemoved") {
        console.error("Subscription was removed - recreating");
        await lifecycle.sync();
      }

      await http.respond(input.request.requestId, {
        statusCode: 202,
        body: { status: "accepted" },
      });
      return;
    }
  }

  await http.respond(input.request.requestId, {
    statusCode: 400,
    body: { error: "Invalid lifecycle notification payload" },
  });
}
