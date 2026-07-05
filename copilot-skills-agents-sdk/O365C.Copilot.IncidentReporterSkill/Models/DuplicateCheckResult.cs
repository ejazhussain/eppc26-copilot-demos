namespace O365C.Copilot.IncidentReporterSkill.Models;

public class DuplicateCheckResult
{
    public string? DuplicateId    { get; set; }
    public string? DuplicateTitle { get; set; }
    public bool    Found          => DuplicateId is not null;
}
