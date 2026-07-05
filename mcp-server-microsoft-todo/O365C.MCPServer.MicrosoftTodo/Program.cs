using Microsoft.Identity.Web;
using O365C.MCPServer.MicrosoftTodo.Services;

var builder = WebApplication.CreateBuilder(args);

//This added to support check user delegate access via WhoAmI Tool
builder.Services.AddHttpContextAccessor();
builder.Services.AddHttpClient();


// Validate incoming Bearer token + OBO exchange + Graph client
builder.Services
    .AddMicrosoftIdentityWebApiAuthentication(builder.Configuration, "AzureAd")
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddMicrosoftGraph(builder.Configuration.GetSection("Graph"))
    .AddInMemoryTokenCaches();


// Validate incoming Bearer tokens from Copilot Studio
//builder.Services
//    .AddMicrosoftIdentityWebApiAuthentication(builder.Configuration, "AzureAd");

builder.Services
    .AddMcpServer() //Registers the MCP server in the DI container
    .WithHttpTransport() //uses Streamable HTTP (the only transport Copilot Studio supports)
    .WithToolsFromAssembly(); //auto-discovers all classes marked [McpServerToolType]


// Register Graph service (scoped � one per request, bound to the request user)
builder.Services.AddScoped<GraphTodoService>();

var app = builder.Build();

app.UseAuthentication();
app.UseAuthorization();

// Require a valid Bearer token on the MCP endpoint
app.MapMcp("/mcp").RequireAuthorization();

app.Run();
