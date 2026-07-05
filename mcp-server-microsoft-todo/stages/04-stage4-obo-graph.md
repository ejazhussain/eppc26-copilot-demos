# Stage 4: OBO Flow + Microsoft Graph (Real To Do Data)

## Outcome

By the end of this stage, you will have:

- A pre-built MCP server with OBO-enabled Graph access running locally
- End-to-end delegated calls from Copilot Studio to Microsoft Graph
- Live Microsoft To Do data flowing through the agent

## Prerequisites

- Stage 3 working (OAuth 2.0 auth, `WhoAmI` returning identity)
- Admin consent granted for `Tasks.ReadWrite` delegated permission in the app registration

---

## How OBO Works (One Sentence)

The server receives a token for your API, exchanges it for a Graph token on behalf of the same user, then calls Graph with user-scoped delegated permissions.

---

## What Was Built (Brief Walk-through)

Rather than live-coding the full Graph integration, switch to the pre-built solution. Briefly show the audience the key pieces:

- **`Program.cs`** — `EnableTokenAcquisitionToCallDownstreamApi()` + `AddMicrosoftGraph(...)` enable the OBO token exchange
- **`Services/GraphTodoService.cs`** — wraps Graph API calls; each method logs the caller's name and UPN to prove delegated user context
- **`Tools/TodoTools.cs`** — 20+ MCP tools backed by `GraphTodoService`
- **`Models/TodoModels.cs`** — simple `TodoList` and `TodoTask` POCOs

Packages used:

| Package | Version |
|---|---|
| ModelContextProtocol.AspNetCore | 1.4.0 |
| Microsoft.Identity.Web | 4.5.0 |
| Microsoft.Identity.Web.GraphServiceClient | 4.5.0 |
| Microsoft.Graph | 5.103.0 |

---

## Step 1: Verify appsettings.Development.json

The pre-built solution needs the same secret from Stage 3. Open:

```text
C:\GitHub\o365c-mcp-servers\O365C.MCPServer.MicrosoftTodo\appsettings.Development.json
```

Confirm it contains:

```json
{
  "AzureAd": {
    "ClientSecret": "YOUR_SECRET_VALUE"
  }
}
```

If missing, copy the secret value from the Stage 3 project's `appsettings.Development.json`.

---

## Step 2: Stop the Stage 3 Server

In the terminal running the Stage 3 server, press `Ctrl+C`.

---

## Step 3: Run the Pre-Built Solution

```powershell
cd C:\GitHub\o365c-mcp-servers\O365C.MCPServer.MicrosoftTodo
```

```powershell
dotnet run --urls http://localhost:5124
```

The ngrok tunnel is already pointing to port 5124, so no changes are needed in Copilot Studio.

---

## Step 4: Update the MCP Tool in Copilot Studio

The server URL hasn't changed but the tool set has expanded. In Copilot Studio:

1. Open your agent → **Tools**
2. Select the existing MCP tool → **Refresh** (or delete and recreate if refresh is not available)
3. Confirm the new tools appear (e.g. `todo_get_my_lists`, `todo_create_task`, etc.)

---

## Step 5: Live Demo — Copilot Studio Prompts

> **Tip:** open **[Microsoft To Do](https://to-do.office.com)** side-by-side in the browser so the audience can see tasks appear and update in real time.
>
> **Starting point:** Microsoft To Do is empty — that's intentional. The demo builds up from nothing, which makes the live updates more dramatic.
>
> **The MCP differentiator:** each prompt below triggers **multiple tool calls** chained together. The agent reasons about intent, resolves task names to IDs, and sequences the calls automatically — something a direct API call cannot do.

Run these prompts in the agent test chat in order:

---

### Top 10 Demo Prompts (Multi-Step Chaining Flow)

```text
1) Who am I?
```
> Single call — sets the scene. Proves delegated identity (name, email, OID) from the Bearer token.

```text
2) Show me all my To Do lists and tell me how many tasks are in each one
```
> Chains: `todo_get_my_lists` → `todo_get_tasks_by_list` for each list. A raw API caller would need to write a loop manually.

```text
3) Create a task called "Follow up with EPPC speakers" in my Tasks list, set it to high importance, and set the due date to tomorrow
```
> Chains: `todo_create_task` → `todo_update_task`. Watch the starred ⭐ task with a due date appear instantly in the To Do app.

```text
4) Create a task called "EPPC Demo Prep" in my Tasks list and add these checklist items: "Prepare slides", "Rehearse flow", "Validate live demo"
```
> Chains: `todo_create_task` → `todo_create_checklist_item` × 3. Four tool calls from one prompt.

```text
5) Add a due date of tomorrow to "EPPC Demo Prep" and also attach this SharePoint file to it: https://ejazhussain.sharepoint.com/sites/dev/Shared%20Documents/Analysis.pdf
```
> Chains: `todo_update_task` (due date) → `todo_create_linked_resource` (SharePoint URL). Two different tool types in one prompt.

```text
6) Show me the full details of "EPPC Demo Prep" including its checklist items and linked resources
```
> Chains: `todo_get_task_by_id` → `todo_get_checklist_items` → `todo_get_linked_resources`. Agent assembles a complete picture — a raw API call would need 3 separate requests.

```text
7) Create a task called "Overdue Review" in my Tasks list, set the due date to two days ago, and set importance to high
```
> Seeds an overdue task and chains create + update in one go.

```text
8) Show me all my overdue tasks and create a follow-up task for each one called "Chase: <original task name>"
```
> Chains: `todo_get_overdue_tasks` → `todo_create_task` for each result. Dynamic multi-call orchestration — impossible to replicate with a single API call.

```text
9) Mark all checklist items on "EPPC Demo Prep" as complete, then mark the task itself as completed
```
> Chains: `todo_get_checklist_items` → `todo_update_checklist_item` × 3 → `todo_complete_task`. Agent sequences dependent calls automatically.

```text
10) Show me everything I completed today across all my lists
```
> Chains: `todo_get_my_lists` → `todo_get_tasks_by_list` per list → filters by `completedDateTime`. Cross-list aggregation from a single natural language request.

---

### Bonus Prompt (strong closer)

```text
In my Tasks list, find or create a task called "EPPC Session Prep", set it to high importance and in progress, add checklist items "Intro", "Security story", and "Live demo", then show me the full task summary including checklist.
```
> 6+ tool calls from one sentence. Ask the audience: *"How many lines of code would this take with a direct API call?"*

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Graph 403 | Ensure `Tasks.ReadWrite` delegated permission has admin consent |
| Empty results | Confirm the test user has To Do lists and tasks |
| OBO failures | Check `ClientSecret` and verify tenant/client IDs match the app registration |
| Tool errors after code update | Restart the app and start a new Copilot Studio test session |
