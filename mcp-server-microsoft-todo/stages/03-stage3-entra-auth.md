# Stage 3: Secure the MCP Endpoint with Entra ID

## Outcome

By the end of this stage, you will have:

- An Entra ID app registration for your MCP API
- JWT Bearer validation enabled in the server
- `/mcp` returning `401` when no valid token is present
- OAuth 2.0 configured in Copilot Studio
- `WhoAmI` tool confirming delegated identity

## Prerequisites

- Stage 2 working (ngrok tunnel active, Copilot Studio agent created)
- Azure portal access to create app registrations

---

## App Registration Reference Values

| Setting | Value |
|---|---|
| Tenant ID | 3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d |
| Client ID | 7cf2eaf0-7954-4b2a-ab9e-2397c11e0dfb |
| Scope | `api://7cf2eaf0-7954-4b2a-ab9e-2397c11e0dfb/access_as_user` |

---

## Step 1: Create the App Registration

In the Azure portal:

1. Go to **App registrations** → **New registration**
2. Name: `o365c-mcp-todo-dev-svc`
3. Supported account type: **Single tenant**
4. Click **Register**

Then configure:

1. **Expose an API** → set the Application ID URI to `api://{clientId}`
2. Add a scope named `access_as_user`
3. Go to **Certificates & secrets** → create a new client secret and copy the value

> Store the secret in `appsettings.Development.json` only — never in `appsettings.json`.

---

## Step 2: Add the Identity Package

```powershell
cd C:\EPPC2026\TH15-MCP-Server\O365C.MCPServer.MicrosoftTodo
```

```powershell
dotnet add package Microsoft.Identity.Web --version 4.5.0
```

---

## Step 3: Update appsettings.json

```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*",
  "AzureAd": {
    "Instance": "https://login.microsoftonline.com/",
    "TenantId": "3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d",
    "ClientId": "7cf2eaf0-7954-4b2a-ab9e-2397c11e0dfb",
    "Scopes": "access_as_user"
  }
}
```

Create (or update) `appsettings.Development.json`:

```json
{
  "AzureAd": {
    "ClientSecret": "YOUR_SECRET_VALUE"
  }
}
```

---

## Step 4: Update Program.cs

Replace the entire contents of `Program.cs` with:

```csharp
using Microsoft.Identity.Web;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddHttpContextAccessor();

builder.Services
    .AddMicrosoftIdentityWebApiAuthentication(builder.Configuration, "AzureAd"); // Configure authentication with Azure AD

builder.Services
     .AddMcpServer()           // Register MCP services
    .WithHttpTransport()      // Enable Streamable HTTP transport
    .WithToolsFromAssembly(); // Discover [McpServerToolType] classes

var app = builder.Build();

app.UseAuthentication();
app.UseAuthorization();

app.MapMcp("/mcp").RequireAuthorization(); // Map MCP endpoint and require authorization

app.Run();
```

Key additions from Stage 2:
- `AddMicrosoftIdentityWebApiAuthentication` validates JWT Bearer tokens
- `UseAuthentication()` / `UseAuthorization()` middleware order matters
- `.RequireAuthorization()` on `MapMcp` enforces auth at the endpoint

---

## Step 5: Verify the 401 Response

Restart the server:

```powershell
cd C:\EPPC2026\TH15-MCP-Server\O365C.MCPServer.MicrosoftTodo
```

```powershell
dotnet run --urls http://localhost:5124
```

Open the ngrok URL directly in a browser:

```text
https://<your-ngrok-domain>/mcp
```

Expected response: `401 Unauthorized` — the endpoint is now protected.

---

## Step 6: Add the WhoAmI Tool

Update `Tools/TodoTools.cs` to add this method (keep `SayHello` in place):

```csharp
using System.Security.Claims;

[McpServerTool, Description("Returns the identity of the currently authenticated user from the Bearer token.")]
public string WhoAmI(IHttpContextAccessor httpContextAccessor)
{
    var claims = httpContextAccessor.HttpContext?.User?.Claims?.ToList();

    var name = claims?.FirstOrDefault(c => c.Type == "name")?.Value
            ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Name)?.Value
            ?? "unknown";

    var email = claims?.FirstOrDefault(c => c.Type == "preferred_username")?.Value
             ?? claims?.FirstOrDefault(c => c.Type == "upn")?.Value
             ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Upn)?.Value
             ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value
             ?? claims?.FirstOrDefault(c => c.Type == "unique_name")?.Value
             ?? "unknown";

    var oid = claims?.FirstOrDefault(c => c.Type == "oid")?.Value
           ?? claims?.FirstOrDefault(c => c.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value
           ?? "unknown";

    return $"Authenticated as: {name} | Email: {email} | OID: {oid}";
}
```

---

## Step 7: Configure OAuth 2.0 in Copilot Studio

First, delete the existing connection in Power Platform:

1. Navigate to **[https://make.powerautomate.com/connections](https://make.powerautomate.com/connections)**
2. Find the existing MCP/HTTP connection and **delete it**

Then, in Copilot Studio, delete the existing MCP tool from your agent and recreate it with **Authentication: OAuth 2.0 Manual**.

| Field | Value |
|---|---|
| Server name | `O365C MCP Microsoft Todo` |
| Server description | `MCP Server for managing Microsoft Todo tasks` |
| Server URL | `https://<your-ngrok-domain>/mcp` |
| Client ID | `7cf2eaf0-7954-4b2a-ab9e-2397c11e0dfb` |
| Client secret | your app secret |
| Authorization URL | `https://login.microsoftonline.com/3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d/oauth2/v2.0/authorize` |
| Token URL | `https://login.microsoftonline.com/3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d/oauth2/v2.0/token` |
| Refresh URL | `https://login.microsoftonline.com/3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d/oauth2/v2.0/token` |
| Scopes | `api://7cf2eaf0-7954-4b2a-ab9e-2397c11e0dfb/access_as_user offline_access` |

> **Important:** copy the redirect URL generated by Copilot Studio and add it to your app registration under **Authentication → Web platform → Redirect URIs**.

---

## Step 8: Test in Chat

In the agent test chat:

```text
Who am I?
```

Expected response:

```text
Authenticated as: <Your Name> | Email: <your@email.com> | OID: <guid>
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| 401 in Copilot Studio | Recreate the connection and sign in again |
| Invalid audience | Verify `ClientId` and scope values match the app registration |
| Consent prompt loops | Confirm the redirect URL is added to the app registration |
| WhoAmI returns unknown fields | Keep the full claim fallback chain as shown |
