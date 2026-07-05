using Microsoft.Azure.Cosmos;
using O365C.Copilot.IncidentReporterSkill.Models;

namespace O365C.Copilot.IncidentReporterSkill.Services;

public sealed class CosmosDbTicketService : ITicketService
{
    private readonly Container _container;

    public CosmosDbTicketService(CosmosClient client, IConfiguration config)
    {
        var dbId = config["CosmosDb:DatabaseId"] ?? "IncidentReporter";
        var containerId = config["CosmosDb:ContainerId"] ?? "Tickets";
        _container = client.GetContainer(dbId, containerId);
    }

    public async Task<IncidentTicket> CreateTicketAsync(
        CreateTicketRequest request, CancellationToken cancellationToken = default)
    {
        var ticket = new IncidentTicket
        {
            Id = $"INC-{Guid.NewGuid().ToString("N")[..6].ToUpper()}",
            Title = request.Title,
            Description = request.Description,
            Category = request.Category,
            Priority = TicketPriority.Normalise(request.Priority),
            AffectedSystem = request.AffectedSystem,
            ReportedBy = request.ReportedBy,
            ReporterEmail = request.ReporterEmail ?? string.Empty,
            ReporterAadId = request.ReporterAadId,
            Department = request.Department,
            UsersAffected = request.UsersAffected,
            WorkaroundAvailable = request.WorkaroundAvailable,
            IsEscalated = request.IsEscalated,
            Status = TicketStatus.Open,
            CreatedAt = DateTime.UtcNow
        };

        await _container.CreateItemAsync(ticket, new PartitionKey(ticket.Id), cancellationToken: cancellationToken);
        return ticket;
    }

    public async Task<IReadOnlyList<IncidentTicket>> GetAllTicketsAsync(
        CancellationToken cancellationToken = default)
    {
        var query = _container.GetItemQueryIterator<IncidentTicket>(
                        "SELECT * FROM c ORDER BY c.createdAt DESC");
        var results = new List<IncidentTicket>();

        while (query.HasMoreResults)
        {
            var page = await query.ReadNextAsync(cancellationToken);
            results.AddRange(page);
        }

        return results;
    }

    public async Task<IncidentTicket?> GetTicketByIdAsync(
        string id, CancellationToken cancellationToken = default)
    {
        try
        {
            var response = await _container.ReadItemAsync<IncidentTicket>(
                               id, new PartitionKey(id), cancellationToken: cancellationToken);
            return response.Resource;
        }
        catch (CosmosException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return null;
        }
    }

    public async Task<IReadOnlyList<IncidentTicket>> GetTicketsByAadIdAsync(
        string aadObjectId, CancellationToken cancellationToken = default)
    {
        var query = _container.GetItemQueryIterator<IncidentTicket>(
            new QueryDefinition(
                "SELECT * FROM c WHERE c.reporterAadId = @aadId AND c.status NOT IN ('Resolved', 'Closed') ORDER BY c.createdAt DESC")
                .WithParameter("@aadId", aadObjectId));

        var results = new List<IncidentTicket>();
        while (query.HasMoreResults)
        {
            var page = await query.ReadNextAsync(cancellationToken: cancellationToken);
            results.AddRange(page);
        }

        return results;
    }

    public async Task<IReadOnlyList<IncidentTicket>> GetTicketsByUserAsync(
        string reportedBy, CancellationToken cancellationToken = default)
    {
        var query = _container.GetItemQueryIterator<IncidentTicket>(
            new QueryDefinition(
                "SELECT * FROM c WHERE c.reportedBy = @user AND c.status NOT IN ('Resolved', 'Closed') ORDER BY c.createdAt DESC")
                .WithParameter("@user", reportedBy));

        var results = new List<IncidentTicket>();
        while (query.HasMoreResults)
        {
            var page = await query.ReadNextAsync(cancellationToken: cancellationToken);
            results.AddRange(page);
        }

        return results;
    }

    public async Task<IncidentTicket> UpdateTicketStatusAsync(
        string id, string status, CancellationToken cancellationToken = default)
    {
        var ticket = await GetTicketByIdAsync(id, cancellationToken)
                     ?? throw new KeyNotFoundException($"Ticket {id} not found.");

        ticket.Status = status;
        ticket.UpdatedAt = DateTime.UtcNow;

        await _container.UpsertItemAsync(ticket, new PartitionKey(ticket.Id), cancellationToken: cancellationToken);
        return ticket;
    }
}
