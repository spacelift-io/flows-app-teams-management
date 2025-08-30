// Subscriptions
export { messagesSubscription } from "./subscriptions/messagesSubscription.ts";
export { handleSubscriptionNotifications } from "../utils/handleSubscriptionNotifications.ts";

// Messages
export { getMessage } from "./messages/getMessage.ts";
export { listReplies } from "./messages/listReplies.ts";

// Channels
export { listChannels } from "./channels/listChannels.ts";
export { listChannelMembers } from "./channels/listChannelMembers.ts";
export { getChannelInfo } from "./channels/getChannelInfo.ts";
export { createChannel } from "./channels/createChannel.ts";
export { archiveChannel } from "./channels/archiveChannel.ts";
export { unarchiveChannel } from "./channels/unarchiveChannel.ts";
export { inviteUsersToChannel } from "./channels/inviteUsersToChannel.ts";
export { removeUserFromChannel } from "./channels/removeUserFromChannel.ts";

// Teams
export { listTeams } from "./teams/listTeams.ts";
export { addTeamMember } from "./teams/addTeamMember.ts";

// Users
export { getUser } from "./users/getUser.ts";
export { listUsers } from "./users/listUsers.ts";

import { messagesSubscription } from "./subscriptions/messagesSubscription.ts";
import { getMessage } from "./messages/getMessage.ts";
import { listReplies } from "./messages/listReplies.ts";
import { listChannels } from "./channels/listChannels.ts";
import { listChannelMembers } from "./channels/listChannelMembers.ts";
import { getChannelInfo } from "./channels/getChannelInfo.ts";
import { createChannel } from "./channels/createChannel.ts";
import { archiveChannel } from "./channels/archiveChannel.ts";
import { unarchiveChannel } from "./channels/unarchiveChannel.ts";
import { inviteUsersToChannel } from "./channels/inviteUsersToChannel.ts";
import { removeUserFromChannel } from "./channels/removeUserFromChannel.ts";
import { listTeams } from "./teams/listTeams.ts";
import { addTeamMember } from "./teams/addTeamMember.ts";
import { getUser } from "./users/getUser.ts";
import { listUsers } from "./users/listUsers.ts";

export const blocks = {
  messagesSubscription,
  getMessage,
  listReplies,
  listTeams,
  listChannels,
  listChannelMembers,
  getChannelInfo,
  createChannel,
  archiveChannel,
  unarchiveChannel,
  addTeamMember,
  inviteUsersToChannel,
  removeUserFromChannel,
  getUser,
  listUsers,
} as const;
