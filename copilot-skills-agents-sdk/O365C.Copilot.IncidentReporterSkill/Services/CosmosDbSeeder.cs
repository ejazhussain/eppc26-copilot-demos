using Microsoft.Azure.Cosmos;
using O365C.Copilot.IncidentReporterSkill.Models;
using System.Text.Json;

namespace O365C.Copilot.IncidentReporterSkill.Services;

/// <summary>
/// Runs once on startup. Creates the Cosmos DB database and container if they don't exist,
/// then seeds sample ITSM tickets from Data/seed-tickets.json if the container is empty.
/// </summary>
public static class CosmosDbSeeder
{
    public static async Task SeedAsync(CosmosClient client, IConfiguration config, ILogger logger)
    {
        var dbId        = config["CosmosDb:DatabaseId"]  ?? "IncidentReporter";
        var containerId = config["CosmosDb:ContainerId"] ?? "Tickets";

        var db = await client.CreateDatabaseIfNotExistsAsync(dbId);
        var containerResponse = await db.Database.CreateContainerIfNotExistsAsync(
            new ContainerProperties { Id = containerId, PartitionKeyPath = "/id" });

        var container = containerResponse.Container;

        // Only seed if the container is empty
        var countQuery = container.GetItemQueryIterator<int>("SELECT VALUE COUNT(1) FROM c");
        var countPage  = await countQuery.ReadNextAsync();
        var count      = countPage.FirstOrDefault();

        if (count > 0)
        {
            logger.LogInformation("CosmosDbSeeder: container has {Count} tickets — skipping seed.", count);
            return;
        }

        // Load seed data from the JSON file
        var seedPath = Path.Combine(AppContext.BaseDirectory, "Data", "seed-tickets.json");
        if (!File.Exists(seedPath))
        {
            logger.LogWarning("CosmosDbSeeder: seed file not found at {Path} — skipping seed.", seedPath);
            return;
        }

        var json    = await File.ReadAllTextAsync(seedPath);
        var records = JsonSerializer.Deserialize<List<SeedRecord>>(json,
                          new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        if (records is null || records.Count == 0)
        {
            logger.LogWarning("CosmosDbSeeder: seed file is empty — skipping seed.");
            return;
        }

        logger.LogInformation("CosmosDbSeeder: seeding {Count} tickets from seed file...", records.Count);

        var now    = DateTime.UtcNow;
        var seeded = 0;

        foreach (var r in records)
        {
            var ticket = new IncidentTicket
            {
                Id             = r.Id,
                Title          = r.Title,
                Description    = r.Description,
                Category       = r.Category,
                Priority       = r.Priority,
                AffectedSystem = r.AffectedSystem,
                ReportedBy     = r.ReportedBy,
                ReporterEmail  = r.ReporterEmail,
                ReporterAadId  = r.ReporterAadId,
                Department     = r.Department,
                AssignedTo     = r.AssignedTo,
                Resolution     = r.Resolution,
                Status         = r.Status,
                CreatedAt      = ParseRelativeDate(r.CreatedAt, now),
                UpdatedAt      = r.UpdatedAt is null ? null : ParseRelativeDate(r.UpdatedAt, now)
            };

            await container.CreateItemAsync(ticket, new PartitionKey(ticket.Id));
            seeded++;
        }

        logger.LogInformation("CosmosDbSeeder: seeded {Count} tickets.", seeded);
    }

    // Parses relative date strings from the JSON file, e.g. "-3d", "-18h", "-30m".
    // Anything else is treated as UtcNow.
    private static DateTime ParseRelativeDate(string? value, DateTime now)
    {
        if (string.IsNullOrWhiteSpace(value) || value == "now")
            return now;

        if (value.EndsWith('d') && int.TryParse(value[..^1], out var days))
            return now.AddDays(days);

        if (value.EndsWith('h') && int.TryParse(value[..^1], out var hours))
            return now.AddHours(hours);

        if (value.EndsWith('m') && int.TryParse(value[..^1], out var minutes))
            return now.AddMinutes(minutes);

        return now;
    }

    private sealed class SeedRecord
    {
        public string  Id             { get; set; } = string.Empty;
        public string  Title          { get; set; } = string.Empty;
        public string  Description    { get; set; } = string.Empty;
        public string  Category       { get; set; } = string.Empty;
        public string  Priority       { get; set; } = string.Empty;
        public string  AffectedSystem { get; set; } = string.Empty;
        public string  ReportedBy     { get; set; } = string.Empty;
        public string  ReporterEmail  { get; set; } = string.Empty;
        public string? ReporterAadId  { get; set; }
        public string? Department     { get; set; }
        public string? AssignedTo     { get; set; }
        public string? Resolution     { get; set; }
        public string  Status         { get; set; } = string.Empty;
        public string  CreatedAt      { get; set; } = "now";
        public string? UpdatedAt      { get; set; }
    }
}
