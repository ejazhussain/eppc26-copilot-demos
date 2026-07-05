namespace O365C.Copilot.IncidentReporterSkill.Models;

// Stored in ITurnState between conversation turns.
// The SDK serialises/deserialises this automatically — state survives across turns.
public class IncidentState
{
    public int Step { get; set; } = 0;
    public string Title { get; set; } = string.Empty;
    public string Priority { get; set; } = string.Empty;
    public string AffectedSystem { get; set; } = string.Empty;
}
