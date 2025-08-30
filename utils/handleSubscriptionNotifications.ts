import { blocks, messaging } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "./teamsClient.ts";

/**
 * Routes incoming webhook notifications to appropriate subscription blocks
 */
export const handleSubscriptionNotifications = async (
  notifications: any[],
  config: Record<string, any>,
) => {
  for (const notification of notifications) {
    const { changeType, resource } = notification;

    // Only handle message notifications
    if (!resource.includes("/messages")) {
      continue;
    }

    // Find interested message subscription blocks
    const interestedBlocks = (
      await blocks.list({
        typeIds: ["messagesSubscription"],
      })
    ).blocks.filter((block) => {
      const { teamId, channelId } = block.config;
      return (
        resource.includes(teamId as string) &&
        (!channelId || resource.includes(channelId as string))
      );
    });

    // Only process if there are interested blocks
    if (interestedBlocks.length === 0) {
      continue; // No one cares about this message, skip
    }

    // Fetch the full message data from API
    try {
      const accessToken = await getAccessToken(config);
      const messageData = await callGraphApi(`/${resource}`, accessToken);

      // Skip system event messages (app installed, member added, etc.)
      if (messageData.body?.content === "<systemEventMessage/>") {
        continue;
      }

      await messaging.sendToBlocks({
        blockIds: interestedBlocks.map((b) => b.id),
        body: { messageData, changeType },
      });
    } catch (error: any) {
      console.error("Failed to fetch message data:", error.message);
      continue; // Skip this notification
    }
  }
};
