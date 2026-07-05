# Stage 04 — Verify HelloSkill in Teams + Confirm Auth

## Outcome

By the end of this stage:

- HelloSkill is confirmed working end-to-end through Teams
- Authentication between Teams → Copilot Studio → Azure Bot → HelloSkill is verified
- The echo response proves the full pipeline is healthy before moving to the IncidentReporter skill

> **Important:** HelloSkill is an echo skill. It will only be invoked when the agent routes to it via the **Echo Test** topic. Phrases like "my laptop won't connect to the VPN" will be handled by the agent's own AI knowledge — not by HelloSkill. That is correct behaviour at this stage.

---

## Step 1: Enable Generative Orchestration

In Copilot Studio:

1. **Settings → Generative AI**
2. Under **Orchestration**, enable **Generative (preview)**
3. Click **Save**

---

## Step 2: Test HelloSkill via Teams

Open the agent in **Teams** (not the Copilot Studio test panel — Teams confirms the full auth chain).

Type phrases that match the **Echo Test** topic description:

```
hello skill
```

```
test the skill
```

```
echo hello
```

Expected: `Echo: **hello skill**` — the response comes back from your local skill via ngrok.

> If you see `Echo: **hello skill**` in Teams, the entire pipeline is working: Teams → Copilot Studio → Azure Bot → ngrok → HelloSkill → back.

---

## Step 3: Confirm Authentication is Working

Check that the response is NOT prefixed with an auth error and is NOT answered by the agent's own AI.

| What you see | Meaning |
|---|---|
| `Echo: **hello skill**` | Auth working, skill invoked correctly |
| Generic AI response | Topic trigger description not matching — check the Echo Test topic in Copilot Studio |
| No response / timeout | ngrok not running or messaging endpoint wrong — go back to Stage 02 |
| `Authentication failed` | Client secret wrong — check `appsettings.Development.json` |

---

## What to Say on Stage

> *"We are testing in Teams rather than the Copilot Studio test panel because Teams adds the real Azure Bot Service authentication layer.
> If the echo comes back, every piece of the pipeline — identity, tunnel, skill routing — is confirmed working.
> This is our green light to swap in the real IncidentReporter skill in Stage 05."*

---

## Stage 04 Complete Checklist

- [ ] Generative orchestration enabled in **Settings → Generative AI**
- [ ] `echo hello` in Teams returns `Echo: **echo hello**`
- [ ] Response comes from HelloSkill, not the agent's own AI
- [ ] No auth errors in the terminal console


