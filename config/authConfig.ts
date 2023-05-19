import { OnBehalfOfCredentialAuthConfig } from "@microsoft/teamsfx";

const OAuthConfig: OnBehalfOfCredentialAuthConfig = {
  authorityHost: process.env.AAD_APP_OAUTH_AUTHORITY_HOST ?? "https://login.microsoftonline.com",
  clientId: process.env.AAD_APP_CLIENT_ID ?? "",
  tenantId: process.env.AAD_APP_TENANT_ID ?? "common",
  clientSecret: process.env.AAD_APP_CLIENT_SECRET ?? "",
};

const ApiEndpoints = {
  apiOmniEndpoint: process.env.API_OMNI_ENDPOINT ?? "",
  apiCopilotEndpoint: process.env.API_COPILOT_ENDPOINT ?? "",
}

const AuthConfig = {
  botId: process.env.BOT_ID ?? "",
  botPassword: process.env.BOT_PASSWORD,
  botDomain: process.env.BOT_DOMAIN,
  botTenantId: process.env.AAD_APP_TENANT_ID,
  botAppType: process.env.BOT_APP_TYPE,
  oboAuthConfig: OAuthConfig,
  connectionName: process.env.CONNECTION_NAME ?? "",
  apiEndpoints: ApiEndpoints,
};

export default AuthConfig;


