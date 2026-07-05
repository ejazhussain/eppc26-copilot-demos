using O365C.Copilot.IncidentReporterSkill.Models;

namespace O365C.Copilot.IncidentReporterSkill.Services;

// Implement this interface to swap in Cosmos DB, ServiceNow, Azure DevOps, or any ITSM backend.
// The skill only depends on this abstraction — no changes needed in the conversation layer.
public interface ITicketService
{
    Task<IncidentTicket> CreateTicketAsync(CreateTicketRequest request, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<IncidentTicket>> GetAllTicketsAsync(CancellationToken cancellationToken = default);

    Task<IncidentTicket?> GetTicketByIdAsync(string id, CancellationToken cancellationToken = default);

    Task<IncidentTicket> UpdateTicketStatusAsync(string id, string status, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<IncidentTicket>> GetTicketsByUserAsync(string reportedBy, CancellationToken cancellationToken = default);

    Task<IReadOnlyList<IncidentTicket>> GetTicketsByAadIdAsync(string aadObjectId, CancellationToken cancellationToken = default);
}
