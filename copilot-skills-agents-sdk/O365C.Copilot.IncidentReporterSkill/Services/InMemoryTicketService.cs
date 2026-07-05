using O365C.Copilot.IncidentReporterSkill.Models;
using System.Collections.Concurrent;

namespace O365C.Copilot.IncidentReporterSkill.Services;

// Dev/demo implementation — all data lives in memory and is lost on restart.
//
// To switch to Cosmos DB:
//   1. Add Microsoft.Azure.Cosmos NuGet package
//   2. Create CosmosDbTicketService : ITicketService
//   3. In Program.cs replace:
//        builder.Services.AddSingleton<ITicketService, InMemoryTicketService>()
//      with:
//        builder.Services.AddSingleton<ITicketService, CosmosDbTicketService>()
//
// Same pattern applies for ServiceNow, Azure DevOps, Jira, or any other ITSM.
public sealed class InMemoryTicketService : ITicketService
{
    private readonly ConcurrentDictionary<string, IncidentTicket> _store = new();
    private int _counter = 1000;

    public Task<IncidentTicket> CreateTicketAsync(CreateTicketRequest request, CancellationToken cancellationToken = default)
    {
        var id = $"INC-{Interlocked.Increment(ref _counter)}";

        var ticket = new IncidentTicket
        {
            Id = id,
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

        _store[id] = ticket;
        return Task.FromResult(ticket);
    }

    public Task<IReadOnlyList<IncidentTicket>> GetAllTicketsAsync(CancellationToken cancellationToken = default)
    {
        IReadOnlyList<IncidentTicket> list = _store.Values
            .OrderByDescending(t => t.CreatedAt)
            .ToList();

        return Task.FromResult(list);
    }

    public Task<IncidentTicket?> GetTicketByIdAsync(string id, CancellationToken cancellationToken = default)
    {
        _store.TryGetValue(id, out var ticket);
        return Task.FromResult(ticket);
    }

    public Task<IReadOnlyList<IncidentTicket>> GetTicketsByAadIdAsync(string aadObjectId, CancellationToken cancellationToken = default)
    {
        IReadOnlyList<IncidentTicket> list = _store.Values
            .Where(t => t.ReporterAadId == aadObjectId
                     && t.Status != TicketStatus.Resolved
                     && t.Status != TicketStatus.Closed)
            .OrderByDescending(t => t.CreatedAt)
            .ToList();

        return Task.FromResult(list);
    }

    public Task<IReadOnlyList<IncidentTicket>> GetTicketsByUserAsync(string reportedBy, CancellationToken cancellationToken = default)
    {
        IReadOnlyList<IncidentTicket> list = _store.Values
            .Where(t => t.ReportedBy.Equals(reportedBy, StringComparison.OrdinalIgnoreCase)
                     && t.Status != TicketStatus.Resolved
                     && t.Status != TicketStatus.Closed)
            .OrderByDescending(t => t.CreatedAt)
            .ToList();

        return Task.FromResult(list);
    }

    public Task<IncidentTicket> UpdateTicketStatusAsync(string id, string status, CancellationToken cancellationToken = default)
    {
        if (!_store.TryGetValue(id, out var ticket))
            throw new KeyNotFoundException($"Ticket {id} not found.");

        ticket.Status = status;
        ticket.UpdatedAt = DateTime.UtcNow;
        return Task.FromResult(ticket);
    }
}
