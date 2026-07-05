namespace O365C.Copilot.IncidentReporterSkill.Models;

public class IncidentTicket
{
    public string Id { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;   // Access | Software | Hardware | Network | Account
    public string Priority { get; set; } = string.Empty;   // Low | Medium | High | Critical
    public string AffectedSystem { get; set; } = string.Empty;
    public string Status { get; set; } = TicketStatus.Open;

    // Reporter — Graph enrichment (Enhancement 3) populates Email, Department
    public string ReportedBy { get; set; } = string.Empty;
    public string ReporterEmail { get; set; } = string.Empty;
    public string? ReporterAadId { get; set; }
    public string? Department { get; set; }

    // Critical escalation — populated when Priority = Critical (Enhancement 2)
    public string? UsersAffected { get; set; }  // Just me | My team | Entire organisation
    public bool? WorkaroundAvailable { get; set; }  // true | false
    public bool IsEscalated { get; set; }  // true when Critical + no workaround

    // Assignment & resolution — populated once a technician picks up the ticket
    public string? AssignedTo { get; set; }
    public string? Resolution { get; set; }

    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime? UpdatedAt { get; set; }
}

public static class TicketStatus
{
    public const string Open = "Open";
    public const string InProgress = "InProgress";
    public const string Resolved = "Resolved";
    public const string Closed = "Closed";
}

public static class TicketPriority
{
    public const string Low = "Low";
    public const string Medium = "Medium";
    public const string High = "High";
    public const string Critical = "Critical";

    public static readonly string[] All = [Low, Medium, High, Critical];

    public static bool IsValid(string value) =>
        All.Contains(value, StringComparer.OrdinalIgnoreCase);

    public static string Normalise(string value) =>
        All.FirstOrDefault(p => p.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? value;
}

public static class TicketCategory
{
    public const string Access = "Access";
    public const string Software = "Software";
    public const string Hardware = "Hardware";
    public const string Network = "Network";
    public const string Account = "Account";
}
