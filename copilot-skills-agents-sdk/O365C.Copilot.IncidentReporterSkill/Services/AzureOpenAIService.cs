using Azure;
using Azure.AI.OpenAI;
using O365C.Copilot.IncidentReporterSkill.Models;
using OpenAI.Chat;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace O365C.Copilot.IncidentReporterSkill.Services;

public sealed class AzureOpenAIService : IAzureOpenAIService
{
    private readonly ChatClient? _chat;

    public AzureOpenAIService(IConfiguration config)
    {
        var endpoint   = config["AzureOpenAI:Endpoint"];
        var apiKey     = config["AzureOpenAI:ApiKey"];
        var deployment = config["AzureOpenAI:DeploymentName"] ?? "gpt-4o";

        if (!string.IsNullOrWhiteSpace(endpoint) && !string.IsNullOrWhiteSpace(apiKey))
        {
            var client = new AzureOpenAIClient(new Uri(endpoint), new AzureKeyCredential(apiKey));
            _chat = client.GetChatClient(deployment);
        }
    }

    public async Task<DuplicateCheckResult> CheckDuplicateAsync(
        string newTitle,
        IReadOnlyList<IncidentTicket> userTickets,
        CancellationToken cancellationToken = default)
    {
        if (_chat is null || !userTickets.Any())
            return new DuplicateCheckResult();

        var ticketList = string.Join("\n", userTickets.Select(
            t => $"- {t.Id}: \"{t.Title}\" (Status: {t.Status}, Priority: {t.Priority})"));

        var userPrompt =
            $"New issue reported: \"{newTitle}\"\n\n" +
            $"User's existing open tickets:\n{ticketList}\n\n" +
            $"Does any existing ticket describe the same issue as the new report?\n" +
            $"Reply with JSON only, no markdown:\n" +
            $"{{\"duplicateId\": \"<ticket id or null>\", \"reason\": \"<one sentence>\"}}";

        var response = await _chat.CompleteChatAsync(
            [
                new SystemChatMessage(
                    "You are an IT support triage assistant. " +
                    "Only flag a duplicate if the issue is clearly the same problem, not just in the same system. " +
                    "Be strict — different symptoms in the same app are different tickets."),
                new UserChatMessage(userPrompt)
            ],
            new ChatCompletionOptions { ResponseFormat = ChatResponseFormat.CreateJsonObjectFormat() },
            cancellationToken);

        var json   = response.Value.Content[0].Text;
        var parsed = JsonSerializer.Deserialize<AiResponse>(json);

        if (parsed?.DuplicateId is null or "null" or "")
            return new DuplicateCheckResult();

        var match = userTickets.FirstOrDefault(t => t.Id == parsed.DuplicateId);
        return new DuplicateCheckResult
        {
            DuplicateId    = parsed.DuplicateId,
            DuplicateTitle = match?.Title ?? parsed.DuplicateId
        };
    }

    private record AiResponse(
        [property: JsonPropertyName("duplicateId")] string? DuplicateId,
        [property: JsonPropertyName("reason")]      string? Reason
    );
}
