import { AppBlock, events } from "@slflows/sdk/v1";
import { callGraphApi, getAccessToken } from "../../utils/teamsClient.ts";
import { userSchema } from "../../schemas/userSchema.ts";

export const getUser: AppBlock = {
  name: "Get User",
  description:
    "Retrieves a user by ID or User Principal Name (email). Returns user profile information.\n\n**Required Permission:** User.Read.All",
  category: "Users",
  inputs: {
    default: {
      name: "Get User",
      description: "Retrieve a user's profile information.",
      config: {
        userId: {
          name: "User ID or Principal Name",
          description:
            "The user's ID (GUID) or User Principal Name (email address, e.g., user@contoso.com)",
          type: "string",
          required: true,
        },
      },
      async onEvent(input) {
        const accessToken = await getAccessToken(input.app.config);
        const userId = input.event.inputConfig.userId;

        const endpoint = `/users/${encodeURIComponent(userId)}?$select=id,userPrincipalName,displayName,givenName,surname,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone,accountEnabled`;

        const user = await callGraphApi(endpoint, accessToken);
        await events.emit(user);
      },
    },
  },
  outputs: {
    default: {
      name: "User",
      description: "User profile information",
      type: userSchema,
    },
  },
};
