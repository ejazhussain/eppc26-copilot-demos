# Stage 01 — Hello World: Build the Echo Skill

## Outcome

By the end of this stage, you will have:

- A running .NET 10 skill on `http://localhost:3978`
- An echo skill that repeats back whatever the user says
- A health check endpoint returning `HelloSkill — running`
- Static files folder structure ready for the manifest

No auth. No Azure Bot. No Copilot Studio yet.

## Demo Flow at a Glance

1. Scaffold project
2. Add Agents SDK packages
3. Wire `Program.cs`
4. Add `HelloSkill.cs`
5. Copy `AspNetExtensions.cs`
6. Add `wwwroot/` static files
7. Run and verify

## Prerequisites

- .NET 10 SDK: run `dotnet --version` → should show `10.0.x`
- Repo cloned at `C:\GitHub\o365c-agents`

---

## Step 1: Scaffold the Project

```cmd
Navigate to this directory
cd C:\EPPC2026\TH31-Copilot-Skill
```        
```powershell
dotnet new webapi -n HelloSkill --framework net10.0
cd HelloSkill
```

```c# 
Build the app - Run the following command

dotnet build
```

Success check:

```text
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

Copy `.gitignore` from `O365C.Copilot.IncidentReporterSkill\.gitignore` into the `HelloSkill` root so secrets never accidentally get committed.

---

## Step 2: Add SDK Packages

```powershell
dotnet add package Microsoft.Agents.Hosting.AspNetCore --version 1.6.55-beta
dotnet add package Microsoft.Agents.Authentication.Msal   --version 1.6.55-beta
```

Verify in `HelloSkill.csproj`:

```powershell
cat HelloSkill.csproj
```

The XML should contain:

```xml
<ItemGroup>
  <PackageReference Include="Microsoft.Agents.Authentication.Msal"  Version="1.6.55-beta" />
  <PackageReference Include="Microsoft.Agents.Hosting.AspNetCore"   Version="1.6.55-beta" />
</ItemGroup>
```

---

## Step 3: Wire Program.cs

Replace the entire contents of `Program.cs`:

```csharp
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddHttpClient();
builder.Logging.AddConsole();

// CORS — required for Copilot Studio test client preflight requests
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .WithMethods("POST", "GET", "OPTIONS");
    });
});

builder.AddAgent<HelloSkill>();

// In-memory state — conversation state survives turns in same session
builder.Services.AddSingleton<IStorage, MemoryStorage>();

// Validate incoming JWT tokens from Azure Bot Service / Copilot Studio
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

var app = builder.Build();

app.UseStaticFiles();   // serves wwwroot/ files
app.UseCors();          // must come BEFORE UseAuthentication
app.UseAuthentication();
app.UseAuthorization();

app.MapAgentRootEndpoint();    // GET / — health check
app.MapAgentApplicationEndpoints(requireAuth: !app.Environment.IsDevelopment());
app.MapControllers();

app.Urls.Add("http://localhost:3978");

app.Run();
```

---

## Step 4: Add HelloSkill.cs

Create `HelloSkill.cs` in the project root:

```csharp
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;

[Agent(name: "HelloSkill", description: "A simple echo skill for Copilot Studio", version: "1.0")]
[AgentInterface(protocol: AgentTransportProtocol.ActivityProtocol, path: "/api/messages")]
public class HelloSkill : AgentApplication
{
    public HelloSkill(AgentApplicationOptions options) : base(options)
    {
        OnConversationUpdate(ConversationUpdateEvents.MembersAdded, WelcomeAsync);
        OnActivity(ActivityTypes.Message, OnMessageAsync, rank: RouteRank.Last);
    }

    private async Task WelcomeAsync(ITurnContext ctx, ITurnState state, CancellationToken ct)
    {
        foreach (var member in ctx.Activity.MembersAdded)
        {
            if (member.Id != ctx.Activity.Recipient.Id)
            {
                await ctx.SendActivityAsync(
                    MessageFactory.Text(
                        "Hello from the skill! Say anything and I'll echo it back.",
                        inputHint: InputHints.ExpectingInput),
                    cancellationToken: ct);
            }
        }
    }

    private async Task OnMessageAsync(ITurnContext ctx, ITurnState state, CancellationToken ct)
    {
        var input = ctx.Activity.Text?.Trim() ?? string.Empty;

        await ctx.SendActivityAsync(
            MessageFactory.Text($"Echo: **{input}**", inputHint: InputHints.IgnoringInput),
            cancellationToken: ct);

        // Tell Copilot Studio the skill is done — control returns to the agent
        await ctx.SendActivityAsync(
            Activity.CreateEndOfConversationActivity(),
            cancellationToken: ct);
    }
}
```

Key rules:

| Rule | Why |
|---|---|
| `[Agent]` and `[AgentInterface]` attributes required | SDK uses them to route inbound activities |
| Always send `EndOfConversationActivity` when done | Without it Copilot Studio hangs waiting |
| State keys must start with `conversation.` | MemoryStorage won't persist without this prefix |

---

## Step 5: Copy AspNetExtensions.cs

`AspNetExtensions.cs` is a helper file that checks every incoming request has a valid token from Azure — think of it as the bouncer at the door. Without it, anyone could call your skill. We copy it as-is; no edits needed.

Copy from the assets folder into your `HelloSkill` root:

```
"C:\EPPC2026\TH31-Copilot-Skill\assets\AspNetExtensions.cs"
```

No changes needed — `AddAgentAspNetAuthentication` in `Program.cs` depends on it.

---

## Step 6: Add Static Files (wwwroot)

Create folders and files:

```
wwwroot/
  icon.png                          ← copy from O365C.Copilot.IncidentReporterSkill\wwwroot\
  privacy.html                      ← copy from O365C.Copilot.IncidentReporterSkill\wwwroot\
  manifest/
    hello-skill-manifest.json       ← placeholder for now; filled at Stage 03
```

Create a blank `hello-skill-manifest.json` for now:

```json
{}
```

Add to `HelloSkill.csproj` so files copy to output:

```xml
<ItemGroup>
  <Content Update="wwwroot\icon.png">
    <CopyToOutputDirectory>Always</CopyToOutputDirectory>
  </Content>
  <Content Update="wwwroot\privacy.html">
    <CopyToOutputDirectory>Always</CopyToOutputDirectory>
  </Content>
  <Content Update="wwwroot\manifest\hello-skill-manifest.json">
    <CopyToOutputDirectory>Always</CopyToOutputDirectory>
  </Content>
</ItemGroup>
```

---

## Step 7: Run and Verify

```powershell
dotnet run
```

Open browser → `http://localhost:3978`

Expected: `HelloSkill — running`

If you see a 404 or error: check `app.MapAgentRootEndpoint()` is in `Program.cs`.

---

## File Summary

```
HelloSkill/
├── HelloSkill.cs                     ← skill class (Step 4)
├── AspNetExtensions.cs               ← copied from IncidentReporterSkill (Step 5)
├── Program.cs                        ← updated in Step 3
├── appsettings.json                  ← updated in Stage 02
├── HelloSkill.csproj                 ← packages added in Step 2
└── wwwroot/
    ├── icon.png
    ├── privacy.html
    └── manifest/
        └── hello-skill-manifest.json ← completed in Stage 03
```

---

## Stage 01 Complete Checklist

- [ ] `dotnet run` starts without errors
- [ ] `http://localhost:3978` returns `HelloSkill — running`
- [ ] `HelloSkill.cs` and `AspNetExtensions.cs` present in project root
- [ ] `wwwroot/` structure created

Next: [Stage 02 — Azure Bot + ngrok](02-azure-bot-ngrok.md)
