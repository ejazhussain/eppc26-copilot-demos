# Stage 02 — Azure Bot + ngrok

## Outcome

By the end of this stage, you will have:

- An Azure Bot resource pointing at your local skill via ngrok
- A client secret in `appsettings.json`
- The skill responding via Azure Bot Web Chat

No Copilot Studio yet — Web Chat confirms the skill + Azure Bot + tunnel pipeline is working before you touch Copilot Studio.

## Stage 02 Flow

1. Open the existing **IncidentReporterSkill** Azure Bot
2. Copy App ID + Tenant ID
3. Reuse (or generate) client secret from existing app registration
4. Set publisher domain
5. Start ngrok
6. Set messaging endpoint in Azure Bot
7. Update `appsettings.json`
8. Test in Azure Bot Web Chat

---

## Step 1: Open the Existing Azure Bot

> We reuse the **IncidentReporterSkill** Azure Bot and its app registration to save time during the demo — no new Azure resources needed.

[portal.azure.com](https://portal.azure.com) → Resource group **rg-copilot-incidentreporterskill** → open **IncidentReporterSkill** (Azure Bot)

---

## Step 2: App ID and Tenant ID

Already known — no need to copy from the portal:

| Value | ID |
|---|---|
| **Microsoft App ID** | `fc20cf97-0114-4644-a3ac-fa09c04d50b6` |
| **App Tenant ID** | `3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d` |

---

## Step 3: Get the Client Secret

The secret is already in the existing project. Open `O365C.Copilot.IncidentReporterSkill\appsettings.json` and copy the `ClientSecret` value from line 22.

Now add it to `HelloSkill\appsettings.Development.json` so it stays out of source control:

```json
{
  "Connections": {
    "ServiceConnection": {
      "Settings": {
        "ClientSecret": "<paste secret here>"
      }
    }
  }
}
```

ASP.NET merges `appsettings.Development.json` on top of `appsettings.json` at runtime — the secret overrides the placeholder without touching the main file.

> `appsettings.Development.json` is already covered by the `.gitignore` copied in Stage 01 Step 1.



## Step 4: Start ngrok

```powershell
ngrok http 3978 --domain=office365clinic.ngrok.dev
```

Confirm tunnel is active:

```text
Forwarding  https://office365clinic.ngrok.dev -> http://localhost:3978
```

Keep this terminal running for the rest of the session.

---

## Step 5: Set Messaging Endpoint in Azure Bot

Back in Azure Bot → **Settings → Configuration**

Set **Messaging endpoint**:
```
https://office365clinic.ngrok.dev/api/messages
```

Click **Apply**.

---

## Step 4: Set Home Page URL (Required before Copilot Studio)

In the app registration → **Branding & properties**

Set **Home page URL** to the manifest file URL:
```
https://office365clinic.ngrok.dev/manifest/helloskill-manifest.json
```

Click **Save**.

> This is a mandatory step. Copilot Studio validates that the manifest's `endpointUrl` domain matches the publisher domain of the app registration. If this URL is missing or wrong you will get `MANIFEST_ENDPOINT_ORIGIN_MISMATCH` when importing the skill in Stage 03.

---

## Step 7: Update appsettings.json

Replace `appsettings.json` with:

```json
{
  "TokenValidation": {
    "Enabled": true,
    "Audiences": [ "fc20cf97-0114-4644-a3ac-fa09c04d50b6" ],
    "TenantId": "3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d"
  },
  "AgentApplication": {
    "StartTypingTimer": false,
    "RemoveRecipientMention": false,
    "NormalizeMentions": false
  },
  "Connections": {
    "ServiceConnection": {
      "Assembly": "Microsoft.Agents.Authentication.Msal",
      "Type": "MsalAuth",
      "Settings": {
        "AuthType": "ClientSecret",
        "AuthorityEndpoint": "https://login.microsoftonline.com/3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d",
        "ClientId": "fc20cf97-0114-4644-a3ac-fa09c04d50b6",
        "ClientSecret": "",
        "TenantId": "3f4d536c-9ebc-4eb1-8304-0f0f2f840b5d",
        "Scopes": [ "https://api.botframework.com/.default" ]
      }
    }
  },
  "ConnectionsMap": [
    { "ServiceUrl": "*", "Connection": "ServiceConnection" }
  ],
  "AllowedCallers": [ "*" ],
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*"
}
```

Only one placeholder to fill in:

| Placeholder | Where to find it |
|---|---|
| `<YOUR_CLIENT_SECRET>` | Copy from `O365C.Copilot.IncidentReporterSkill\appsettings.json` → `Connections.ServiceConnection.Settings.ClientSecret` |

> `AllowedCallers: ["*"]` is acceptable for development. You will replace `"*"` with the Copilot Studio Agent App ID in Stage 03.

---

## Step 8: Test in Azure Bot Web Chat

Restart the skill so the new `appsettings.json` is loaded:

```powershell
dotnet run
```

In the Azure Bot resource → **Test in Web Chat**

Type: `hello`

Expected: `Echo: **hello**`

If no response, check the [Troubleshooting](#troubleshooting) table below before proceeding to Stage 03.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| No response at all | Messaging endpoint wrong or ngrok not running | Check endpoint URL; verify ngrok is forwarding to port 3978 |
| `Authentication failed` | Client secret wrong or expired | Re-check `ClientSecret` in appsettings |
| `Sorry, my bot code is having issues` | Skill threw an unhandled exception | Check the Visual Studio / terminal console for the error |

---

## Stage 02 Complete Checklist

- [ ] Azure Bot resource created
- [ ] ngrok tunnel active: `https://office365clinic.ngrok.dev -> http://localhost:3978`
- [ ] `appsettings.json` has real Client ID, Tenant ID, and Secret
- [ ] Azure Bot Web Chat returns `Echo: **hello**`

Next: [Stage 03 — Copilot Studio Integration](03-copilot-studio.md)
