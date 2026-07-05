using O365C.Copilot.IncidentReporterSkill.Models;

namespace O365C.Copilot.IncidentReporterSkill.Services;

public interface IAzureOpenAIService
{
    Task<DuplicateCheckResult> CheckDuplicateAsync(
        string newTitle,
        IReadOnlyList<IncidentTicket> userTickets,
        CancellationToken cancellationToken = default);
}
