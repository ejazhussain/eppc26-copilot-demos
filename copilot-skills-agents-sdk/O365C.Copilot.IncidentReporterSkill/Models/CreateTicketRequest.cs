namespace O365C.Copilot.IncidentReporterSkill.Models;

public class CreateTicketRequest
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string Priority { get; set; } = string.Empty;
    public string AffectedSystem { get; set; } = string.Empty;
    public string ReportedBy { get; set; } = string.Empty;

    // Populated by Graph enrichment (Enhancement 3) — optional at creation time
    public string? ReporterEmail { get; set; }
    public string? ReporterAadId { get; set; }
    public string? Department { get; set; }

    // Critical escalation — populated when Priority = Critical (Enhancement 2)
    public string? UsersAffected { get; set; }
    public bool? WorkaroundAvailable { get; set; }
    public bool IsEscalated { get; set; }
}
