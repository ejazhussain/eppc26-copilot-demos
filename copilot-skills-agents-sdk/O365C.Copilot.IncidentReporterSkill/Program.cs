using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using Microsoft.Agents.Storage.Blobs;
using Microsoft.Azure.Cosmos;
using O365C.Copilot.IncidentReporterSkill;
using O365C.Copilot.IncidentReporterSkill.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddHttpClient();
builder.Services.AddOpenApi();
builder.Logging.AddConsole();

// ── CORS ─────────────────────────────────────────────────────────────────────
// Fixes 405 Method Not Allowed on OPTIONS preflight requests from the
// Copilot Studio test client. Must be registered before builder.Build().
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .WithMethods("POST", "GET", "OPTIONS");
    });
});

// ── Ticket service ────────────────────────────────────────────────────────────
var cosmosConnectionString = builder.Configuration["CosmosDb:ConnectionString"];

if (!string.IsNullOrWhiteSpace(cosmosConnectionString))
{
    // Cosmos DB — camelCase serialisation so "Id" maps to the required "id" field
    builder.Services.AddSingleton(_ => new CosmosClient(
        cosmosConnectionString,
        new CosmosClientOptions
        {
            SerializerOptions = new CosmosSerializationOptions
            {
                PropertyNamingPolicy = CosmosPropertyNamingPolicy.CamelCase
            }
        }));

    builder.Services.AddSingleton<ITicketService, CosmosDbTicketService>();
}
else
{
    // Fallback — in-memory, data lost on restart (dev / CI)
    builder.Services.AddSingleton<ITicketService, InMemoryTicketService>();
}

// ── Azure OpenAI ──────────────────────────────────────────────────────────────
builder.Services.AddSingleton<IAzureOpenAIService, AzureOpenAIService>();

// ── Microsoft Graph (app-only user enrichment) ────────────────────────────────
builder.Services.AddSingleton<IGraphUserService, GraphUserService>();

// ── Agents SDK ────────────────────────────────────────────────────────────────
builder.AddAgent<IncidentReporterSkill>();

// ── Conversation State Storage ─────────────────────────────────────────────────
// Dev/CI:    MemoryStorage — fast, zero-config, state lost on restart.
// Production: BlobsStorage — state survives restarts, crashes, and scale-out.
// Set Storage:BlobStorage:ConnectionString in appsettings to enable BlobsStorage.
// The skill never touches IStorage directly — swap here, zero changes to skill code.
var blobConnectionString = builder.Configuration["Storage:BlobStorage:ConnectionString"];
var blobContainer        = builder.Configuration["Storage:BlobStorage:ContainerName"]
                           ?? "incident-reporter-state";

if (!string.IsNullOrWhiteSpace(blobConnectionString))
{
    builder.Services.AddSingleton<IStorage>(
        new BlobsStorage(blobConnectionString, blobContainer));
}
else
{
    // Fallback — in-memory, state lost on restart (dev / CI)
    builder.Services.AddSingleton<IStorage, MemoryStorage>();
}

builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

var app = builder.Build();

// ── Cosmos DB seed ────────────────────────────────────────────────────────────
// Creates the database + container if missing, then seeds sample tickets if empty.
if (!string.IsNullOrWhiteSpace(cosmosConnectionString))
{
    var cosmos = app.Services.GetRequiredService<CosmosClient>();
    var logger = app.Services.GetRequiredService<ILoggerFactory>().CreateLogger("CosmosDbSeeder");
    await CosmosDbSeeder.SeedAsync(cosmos, builder.Configuration, logger);
}

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

// Serves wwwroot/manifest/*.json and wwwroot/icon.png
app.UseStaticFiles();

// ── CORS — must come BEFORE UseAuthentication ─────────────────────────────────
app.UseCors();

app.UseAuthentication();
app.UseAuthorization();

// Health check endpoint
app.MapAgentRootEndpoint();

app.MapAgentApplicationEndpoints(requireAuth: !app.Environment.IsDevelopment());

app.MapControllers();

app.Urls.Add("http://localhost:3978");

app.Run();
