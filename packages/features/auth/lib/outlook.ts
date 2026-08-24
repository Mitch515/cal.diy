import process from "node:process";
export const OUTLOOK_CLIENT_ID = process.env.MS_GRAPH_CLIENT_ID;
export const OUTLOOK_CLIENT_SECRET = process.env.MS_GRAPH_CLIENT_SECRET;
const WEALTH_NAVIGATOR_TENANT_ID = "f1a52f59-393f-43d3-813f-fa1197512059";
export const getOutlookTenantId = () => {
  const configuredTenantId = process.env.MS_GRAPH_TENANT_ID;
  return configuredTenantId && /^[a-zA-Z0-9.-]+$/.test(configuredTenantId)
    ? configuredTenantId
    : WEALTH_NAVIGATOR_TENANT_ID;
};
export const OUTLOOK_TENANT_ID = getOutlookTenantId();
export const OUTLOOK_LOGIN_ENABLED = process.env.OUTLOOK_LOGIN_ENABLED === "true";
export const IS_OUTLOOK_LOGIN_ENABLED = !!(
  OUTLOOK_CLIENT_ID &&
  OUTLOOK_CLIENT_SECRET &&
  process.env.OUTLOOK_LOGIN_ENABLED === "true"
);
