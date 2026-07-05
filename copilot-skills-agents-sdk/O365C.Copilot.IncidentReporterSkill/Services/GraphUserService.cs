using Azure.Identity;
using Microsoft.Graph;

namespace O365C.Copilot.IncidentReporterSkill.Services;

/// <summary>
/// App-only Microsoft Graph client. Fetches user profile fields (displayName, mail, department)
/// using the AAD object ID already present on every incoming activity.
///
/// Requires application permission: User.Read.All (admin-consented).
///
/// Configuration keys (MicrosoftGraph section):
///   TenantId     — AAD tenant GUID
///   ClientId     — App registration client ID
///   ClientSecret — Client secret value
///
/// If any key is missing the service returns null on every call (silent no-op).
/// All exceptions are caught and logged as warnings — the conversation continues unaffected.
/// </summary>
public sealed class GraphUserService : IGraphUserService
{
    private readonly GraphServiceClient? _client;
    private readonly ILogger<GraphUserService> _logger;

    public GraphUserService(IConfiguration config, ILogger<GraphUserService> logger)
    {
        _logger = logger;

        var tenantId = config["MicrosoftGraph:TenantId"];
        var clientId = config["MicrosoftGraph:ClientId"];
        var clientSecret = config["MicrosoftGraph:ClientSecret"];

        if (!string.IsNullOrWhiteSpace(tenantId) &&
            !string.IsNullOrWhiteSpace(clientId) &&
            !string.IsNullOrWhiteSpace(clientSecret))
        {
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            _client = new GraphServiceClient(credential,
                ["https://graph.microsoft.com/.default"]);

            _logger.LogInformation("[GRAPH] GraphUserService initialised for tenant {TenantId}", tenantId);
        }
        else
        {
            _logger.LogInformation("[GRAPH] MicrosoftGraph config not set — enrichment disabled");
        }
    }

    public async Task<GraphUserInfo?> GetUserAsync(string aadObjectId, CancellationToken ct = default)
    {
        if (_client is null || string.IsNullOrWhiteSpace(aadObjectId))
            return null;

        try
        {
            var user = await _client.Users[aadObjectId].GetAsync(req =>
            {
                req.QueryParameters.Select = ["displayName", "mail", "department"];
            }, ct);

            if (user is null)
                return null;

            return new GraphUserInfo(
                user.DisplayName ?? string.Empty,
                user.Mail ?? string.Empty,
                user.Department);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "[GRAPH] Failed to fetch user {AadId} — continuing without enrichment",
                aadObjectId);
            return null;
        }
    }
}
