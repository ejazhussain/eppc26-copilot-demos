namespace O365C.Copilot.IncidentReporterSkill.Services;

public interface IGraphUserService
{
    /// <summary>
    /// Fetches display name, email and department for the given AAD object ID.
    /// Returns null silently if Graph is unavailable or not configured.
    /// </summary>
    Task<GraphUserInfo?> GetUserAsync(string aadObjectId, CancellationToken ct = default);
}

public sealed record GraphUserInfo(
    string DisplayName,
    string Email,
    string? Department);
