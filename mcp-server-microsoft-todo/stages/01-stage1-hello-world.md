# Stage 1: Hello World MCP Server

## Outcome

By the end of this stage, you will have:

- A running .NET 10 MCP server
- One hardcoded tool named `SayHello`
- A working endpoint at `http://localhost:<port>/mcp`
- A successful tool call in MCP Inspector

No auth, no Graph, no external APIs — just the MCP plumbing.

## Prerequisites

- .NET 10 SDK installed
- Node.js installed (for MCP Inspector via `npx`)
- Working directory: `C:\EPPC2026\TH15-MCP-Server`

## Key Files

| File | Purpose |
|---|---|
| `O365C.MCPServer.MicrosoftTodo.csproj` | MCP NuGet reference |
| `Program.cs` | Registers MCP server and maps endpoint |
| `Tools/TodoTools.cs` | First callable MCP tool |

---

## Step 1: Scaffold the Project

```powershell
cd C:\EPPC2026\TH15-MCP-Server
```

```powershell
dotnet new web -n O365C.MCPServer.MicrosoftTodo -f net10.0 --force
```

```powershell
cd O365C.MCPServer.MicrosoftTodo
```

```powershell
dotnet build
```

> `--force` is needed because the folder already contains documentation files.

Expected output:

```text
Build succeeded.  0 Warning(s)  0 Error(s)
```

---

## Step 2: Fix the Port

The scaffolded project picks a random port. Create `Properties/launchSettings.json` to lock it to `5124`:

```json
{
  "$schema": "https://json.schemastore.org/launchsettings.json",
  "profiles": {
    "http": {
      "commandName": "Project",
      "dotnetRunMessages": true,
      "launchBrowser": false,
      "applicationUrl": "http://localhost:5124",
      "environmentVariables": {
        "ASPNETCORE_ENVIRONMENT": "Development"
      }
    }
  }
}
```

> `launchBrowser` is set to `false` — no browser window needed for an API server.

---

## Step 3: Add MCP SDK Package

```powershell
dotnet add package ModelContextProtocol.AspNetCore --version 1.4.0
```

---

## Step 4: Wire MCP in Program.cs

Replace the entire contents of `Program.cs` with:

```csharp
var builder = WebApplication.CreateBuilder(args);

builder.Services
    .AddMcpServer()           // Register MCP services
    .WithHttpTransport()      // Enable Streamable HTTP transport
    .WithToolsFromAssembly(); // Discover [McpServerToolType] classes

var app = builder.Build();

app.MapMcp("/mcp"); // Expose the MCP endpoint

app.Run();
```

---

## Step 5: Add Your First Tool

Create `Tools/TodoTools.cs` **inside the project folder** (`O365C.MCPServer.MicrosoftTodo\Tools\TodoTools.cs`):

```csharp
using ModelContextProtocol.Server;
using System.ComponentModel;

namespace O365C.MCPServer.MicrosoftTodo.Tools;

[McpServerToolType]  // Marks this class as a tool container
public class TodoTools
{
    [McpServerTool, Description("Returns a greeting message from the O365C MicrosoftTodo MCP server.")]
    public string SayHello()
    {
        return "Hello from O365C.MCPServer.MicrosoftTodo! The server is running correctly.";
    }
}
```

**Attribute quick reference:**

| Attribute | Purpose |
|---|---|
| `[McpServerToolType]` | Marks the class as a tool container |
| `[McpServerTool]` | Marks the method as callable by MCP clients |
| `[Description]` | Helps AI clients decide when to call the tool |

---

## Step 6: Run and Test with MCP Inspector

**Terminal 1 — start the server:**

```powershell
dotnet run --urls http://localhost:5124
```

Look for:

```text
Now listening on: http://localhost:5124
```

Leave this terminal running.

**Terminal 2 — start MCP Inspector:**

```powershell
npx @modelcontextprotocol/inspector
```

Open the URL shown (usually `http://localhost:5173`).

**In the Inspector UI:**

1. Set **Transport Type** → `Streamable HTTP`
2. Set **URL** → `http://localhost:5124/mcp`
3. Click **Connect** — expect `Connected`
4. Open **Tools** → select `SayHello` → click **Run Tool**

Expected response:

```text
Hello from O365C.MCPServer.MicrosoftTodo! The server is running correctly.
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Inspector won't connect | Confirm URL ends with `/mcp` |
| `SayHello` not listed | Confirm `[McpServerToolType]` on the class and `.WithToolsFromAssembly()` in `Program.cs` |
| Port is not 5124 | Use whatever port `dotnet run` prints |
