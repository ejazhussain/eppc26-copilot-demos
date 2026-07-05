# Stage 2: Connect to Copilot Studio via ngrok

## Outcome

By the end of this stage, you will have:

- A public HTTPS endpoint forwarding to your local MCP server
- A Copilot Studio agent connected to your MCP endpoint
- A successful tool call from Copilot Studio to your local app

No code changes are required in this stage.

## Prerequisites

- Stage 1 working locally
- ngrok installed and authenticated
- Copilot Studio access available

---

## Step 1: Run the Local Server

```powershell
cd C:\EPPC2026\TH15-MCP-Server\O365C.MCPServer.MicrosoftTodo
```

```powershell
dotnet run --urls http://localhost:5124
```

Expected output:

```text
Now listening on: http://localhost:5124
```

Leave this terminal running.

---

## Step 2: Start ngrok Tunnel

Open a second terminal.

**Free tier** (random URL — changes on every restart):

```powershell
ngrok http 5124
```

**Static domain** (recommended for demos — URL survives restarts):

```powershell
ngrok http --domain=office365clinic.ngrok.dev 5124
```

Expected output:

```text
Forwarding  https://<your-ngrok-domain> -> http://localhost:5124
```

Your MCP endpoint is:

```text
https://<your-ngrok-domain>/mcp
```

---

## Step 3: Create or Open a Copilot Studio Agent

1. Open [https://copilotstudio.microsoft.com](https://copilotstudio.microsoft.com)
2. Create a new agent (or open an existing one)
3. Use a descriptive name, e.g. **Microsoft Todo Agent**

Suggested agent instructions:

```text
You are Microsoft Todo Assistant.
Use MCP tools for all Microsoft To Do data operations.
Do not invent tasks, list IDs, or results.
If required inputs are missing or ambiguous, ask a focused clarifying question.
After write operations, confirm exactly what changed.
```

> For the full production instruction set, see `Documentation/07-copilot-studio-agent-instructions.md`.

---

## Step 4: Add the MCP Server Tool

In agent **Tools**:

1. Click **Add a tool**
2. Select **Model Context Protocol**
3. Fill in the values:

| Field | Value |
|---|---|
| Server name | O365C MicrosoftTodo MCP |
| Description | Manages Microsoft To Do tasks and lists |
| Server URL | `https://<your-ngrok-domain>/mcp` |
| Authentication | None |

4. Click **Create**

Expected result: Copilot Studio connects and `SayHello` appears in the tool list.

---

## Step 5: Create Connection and Enable Tool

1. On the connection row, select **Create new connection**
2. Since auth is set to None, the connection completes immediately
3. Click **Add and configure**

Expected result: tool is enabled and a confirmation banner appears.

---

## Step 6: Test in Chat

In the agent test chat:

```text
Say hello
```

Expected response:

```text
Hello from O365C.MCPServer.MicrosoftTodo! The server is running correctly.
```

Check your local server terminal — you should see an incoming request log.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Connect fails | Confirm URL ends with `/mcp` |
| Timeout | Confirm `dotnet run` is still running on port 5124 |
| URL changed after restart | ngrok free-tier URL changes each run — update the Copilot Studio connection |
| Tool not visible | Recreate the MCP tool or start a new test session |
| Chat did not call tool | Use an explicit prompt: *Use the say hello tool* |
