# Stage 03 — Copilot Studio Integration

## Outcome

By the end of this stage, you will have:

- A skill manifest served at your public ngrok URL
- The skill registered in Copilot Studio with status Active
- A topic that routes conversations to the skill
- A passing end-to-end echo test via the Copilot Studio test panel

## Stage 03 Flow

1. Create the skill manifest
2. Verify all three public URLs respond correctly
3. Create the Copilot Studio agent
4. Get the Copilot Studio App ID → update `AllowedCallers`
5. Register the skill via manifest URL
6. Create a topic with trigger phrases
7. Publish and test

---

## Step 1: Create the Skill Manifest

Replace the placeholder `wwwroot/manifest/hello-skill-manifest.json` with:

```json
{
  "$schema": "https://schemas.botframework.com/schemas/skills/skill-manifest-2.0.0.json",
  "$id": "HelloSkill",
  "name": "Hello Skill",
  "version": "1.0",
  "description": "A simple echo skill — repeats back whatever the user says. Use this to verify the skill pipeline is working end to end.",
  "publisherName": "Office365 Clinic",
  "privacyUrl": "https://office365clinic.ngrok.dev/privacy.html",
  "copyright": "Copyright 2026 Office365 Clinic",
  "iconUrl": "https://office365clinic.ngrok.dev/icon.png",
  "tags": [ "echo", "hello", "demo" ],
  "endpoints": [
    {
      "name": "default",
      "protocol": "BotFrameworkV3",
      "description": "Default endpoint",
      "endpointUrl": "https://office365clinic.ngrok.dev/api/messages",
      "msAppId": "<YOUR_CLIENT_ID>"
    }
  ],
  "activities": {
    "handleRequest": {
      "description": "Echoes the user's message back to them",
      "type": "message"
    }
  }
}
```

Replace `<YOUR_CLIENT_ID>` with the Azure Bot Microsoft App ID from Stage 02.

---

## Step 2: Verify All Three Public URLs

Rebuild and restart the skill, then confirm all three return expected content:

| URL | Expected |
|---|---|
| `https://office365clinic.ngrok.dev` | `HelloSkill — running` |
| `https://office365clinic.ngrok.dev/manifest/helloskill-manifest.json` | Raw JSON |
| `https://office365clinic.ngrok.dev/icon.png` | Icon image renders |

Common fixes:
- Manifest returns 404 → folder is named `wwroot` not `wwwroot`, or `CopyToOutputDirectory` is missing in `.csproj`
- Health check fails → skill not running or ngrok not started on port 3978

---

## Step 3: Create Copilot Studio Agent

Open [copilotstudio.microsoft.com](https://copilotstudio.microsoft.com) → **+ Create → New agent → Skip to configure**

| Field | Value |
|---|---|
| Name | `IT Helpdesk` |
| Description | `Helps employees log and track IT support requests` |
| Instructions | `You are an IT Helpdesk assistant. Help users report technical issues, check ticket status, and escalate when needed. Be concise and professional.` |

Click **Create**.

---

## Step 4: Get the Copilot Studio App ID

In your agent → **Settings → Skills → Add a skill**

The dialog shows:

```
To allow this skill to identify your agent, add the following App ID to the skill's allow list:

  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx   ← Copy this
```

Copy the App ID. Update `appsettings.json`:

```json
"AllowedCallers": [ "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" ]
```

Restart the skill.

---

## Step 5: Register Skill via Manifest URL

Back in the **Add a skill** dialog, paste:

```
https://office365clinic.ngrok.dev/manifest/helloskill-manifest.json
```

Click **Next** — Copilot Studio fetches and validates the manifest.

**Copilot Studio validation checks:**

| Check | Pass condition |
|---|---|
| Manifest reachable | HTTP 200 with valid JSON |
| `name` present | Non-empty string |
| `msAppId` | Matches a registered Azure Bot |
| `endpointUrl` domain | Matches the Azure Bot publisher domain |
| `activities[*].description` | Non-empty for each activity |

If validation passes → click **Add skill**

The skill appears in **Settings → Skills** with status **Active**.

---

## Step 6: Create a Topic

**Topics → + Add topic → From blank**

- Name: `Echo Test`
- In **"Describe what the topic does"**:
  ```
  This topic lets the user test the echo skill. Use it when the user wants to echo a message, test the skill pipeline, or say hello to the skill.
  ```
- Delete the default Message node
- Click **+** → **Call an action** → select **Hello Skill → handleRequest**
- Click **Save** → **Publish**

> The description is what the AI reads to decide when to route to this topic — make it clear and specific.

---

## Step 7: Test End-to-End

Click **Test** (top right in Copilot Studio) → type: `echo this`

Expected flow:

| Turn | Speaker | Message |
|---|---|---|
| 1 | Copilot | Routes to skill |
| 2 | Skill | `Hello from the skill! Say anything and I'll echo it back.` |
| 3 | You | `Hello conference audience` |
| 4 | Skill | `Echo: **Hello conference audience**` |
| 5 | Copilot | Control returns — agent responds |

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `MANIFEST_ENDPOINT_ORIGIN_MISMATCH` | `endpointUrl` domain ≠ Azure Bot publisher domain | Set Home page URL in app registration to match ngrok domain (Stage 02 Step 4) |
| `MANIFEST_MALFORMED` | Required field missing or empty | Check `name`, `msAppId`, `endpointUrl`, `activities[*].description` are all set |
| `401 Unauthorized` on `/api/messages` | Copilot Studio App ID not in `AllowedCallers` | Add it to `appsettings.json` and restart |
| Skill times out / no response | ngrok not running or skill not started | Check health endpoint; verify `dotnet run` is active |
| Skill registered but missing from "Call an action" | Sync delay after registration | Close and reopen the topic editor; or remove and re-add the skill |
| Copilot Studio hangs after skill responds | Missing `EndOfConversationActivity` | Verify `Activity.CreateEndOfConversationActivity()` is sent at end of `OnMessageAsync` |

---

## Stage 03 Complete Checklist

- [ ] Manifest JSON accessible at public URL
- [ ] Skill status **Active** in Copilot Studio Settings → Skills
- [ ] `AllowedCallers` contains the Copilot Studio App ID (not `"*"`)
- [ ] Agent published
- [ ] Echo test passes — `Echo: **Hello conference audience**` returned

Next: [Stage 04 — Generative Orchestration](04-generative-orchestration.md)
