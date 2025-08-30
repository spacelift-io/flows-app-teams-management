/**
 * Shared JSON Schema for Microsoft Teams user objects
 * Used by both getUser and listUsers blocks
 */
export const userSchema = {
  type: "object",
  description: "Microsoft Teams user",
  properties: {
    id: { type: "string", description: "User ID" },
    userPrincipalName: {
      type: "string",
      description: "User Principal Name (email)",
    },
    displayName: { type: "string", description: "Display name" },
    givenName: { type: "string", description: "First name" },
    surname: { type: "string", description: "Last name" },
    mail: { type: "string", description: "Email address" },
    jobTitle: { type: "string", description: "Job title" },
    department: { type: "string", description: "Department" },
    officeLocation: { type: "string", description: "Office location" },
    businessPhones: {
      type: "array",
      description: "Business phone numbers",
      items: { type: "string" },
    },
    mobilePhone: { type: "string", description: "Mobile phone" },
    accountEnabled: {
      type: "boolean",
      description: "Whether the account is enabled",
    },
  },
  required: ["id", "userPrincipalName", "displayName"],
};
