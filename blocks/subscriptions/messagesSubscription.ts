import { AppBlock, events } from "@slflows/sdk/v1";
import { messageEventSchema } from "../../schemas/messageSchema.ts";

export const messagesSubscription: AppBlock = {
  name: "Messages Subscription",
  description:
    "Subscribes to message events in Teams channels (created, updated, deleted). Includes reactions as 'updated' events. The app automatically manages the webhook subscription.\n\n**Required Permission:** ChannelMessage.Read.All",
  category: "Subscriptions",
  config: {
    teamId: {
      name: "Team ID",
      description: "The ID of the team to monitor for messages.",
      type: "string",
      required: true,
    },
    channelId: {
      name: "Channel ID (Optional)",
      description:
        "If specified, only messages from this specific channel will be received. Leave empty to receive messages from all channels in the team.",
      type: "string",
      required: false,
    },
  },
  async onInternalMessage({ message }) {
    const { messageData, changeType } = message.body;
    await events.emit({ ...messageData, changeType });
  },
  outputs: {
    default: {
      name: "On Message",
      description: "Emitted when a message event occurs in a Teams channel.",
      type: messageEventSchema,
    },
  },
};
