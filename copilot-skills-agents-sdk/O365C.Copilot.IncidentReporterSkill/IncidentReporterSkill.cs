using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using O365C.Copilot.IncidentReporterSkill.Models;
using O365C.Copilot.IncidentReporterSkill.Services;

namespace O365C.Copilot.IncidentReporterSkill;

[Agent(name: "IncidentReporterSkill", description: "IT Incident Reporter skill", version: "1.0")]
[AgentInterface(protocol: AgentTransportProtocol.ActivityProtocol, path: "/api/messages")]
public class IncidentReporterSkill : AgentApplication
{
    // ── Conversation steps ───────────────────────────────────────────────────
    // 0 — collect title          → run duplicate check
    // 1 — duplicate confirmation → only reached when a duplicate was found
    // 2 — collect description
    // 3 — collect priority
    //       Critical → 4          Other → 6   (adaptive branch)
    // 4 — [Critical only] how many users affected?
    // 5 — [Critical only] is there a workaround?
    // 6 — collect affected system → build confirmation summary
    // 7 — confirmation summary   → Yes → create ticket / No → cancel
    private const string StepKey = "conversation.incident_step";
    private const string TitleKey = "conversation.incident_title";
    private const string DescriptionKey = "conversation.incident_description";
    private const string PriorityKey = "conversation.incident_priority";
    private const string SystemKey = "conversation.incident_system";
    private const string DupIdKey = "conversation.incident_dup_id";
    private const string DupTitleKey = "conversation.incident_dup_title";
    private const string UsersAffectedKey = "conversation.incident_users_affected";
    private const string WorkaroundKey = "conversation.incident_workaround";
    private const string ReporterEmailKey = "conversation.incident_reporter_email";
    private const string DepartmentKey = "conversation.incident_department";

    private readonly ITicketService _tickets;
    private readonly IAzureOpenAIService _openAI;
    private readonly IGraphUserService _graph;
    private readonly string? _fallbackAadId;

    public IncidentReporterSkill(
        AgentApplicationOptions options,
        ITicketService tickets,
        IAzureOpenAIService openAI,
        IGraphUserService graph,
        IConfiguration config)
        : base(options)
    {
        _tickets = tickets;
        _openAI = openAI;
        _graph = graph;
        _fallbackAadId = config["Demo:FallbackAadObjectId"];

        OnConversationUpdate(ConversationUpdateEvents.MembersAdded, WelcomeAsync);
        OnActivity(ActivityTypes.EndOfConversation, OnEndOfConversationAsync);
        OnActivity(ActivityTypes.Message, OnMessageAsync, rank: RouteRank.Last);
    }

    // ── Welcome ──────────────────────────────────────────────────────────────

    private async Task WelcomeAsync(ITurnContext ctx, ITurnState state, CancellationToken ct)
    {
        foreach (var member in ctx.Activity.MembersAdded)
        {
            if (member.Id != ctx.Activity.Recipient.Id)
            {
                await ctx.SendActivityAsync(
                    MessageFactory.Text(
                        "Hi! I'm the **IT Incident Reporter**.\n\n" +
                        "Please briefly summarise the issue in one line.",
                        inputHint: InputHints.ExpectingInput),
                    cancellationToken: ct);
            }
        }
    }

    // ── Message handler ──────────────────────────────────────────────────────

    private async Task OnMessageAsync(ITurnContext ctx, ITurnState state, CancellationToken ct)
    {
        var step = GetInt(state, StepKey);
        var input = ctx.Activity.Text?.Trim() ?? string.Empty;

        // Copilot Studio does not set Activity.Name — activityAction in Activity.Value
        // is the explicit routing signal passed from the topic's skill action node.
        if (step == 0 &&
            ctx.Activity.Value is System.Text.Json.JsonElement json &&
            json.TryGetProperty("activityAction", out var actionProp))
        {
            switch (actionProp.GetString())
            {
                case "checkTicketStatus":
                    await OnCheckTicketStatusAsync(ctx, state, ct);
                    return;
            }
        }

        Console.WriteLine($"[SKILL] Step={step}, Input={input}");

        try
        {
            switch (step)
            {
                // ── Step 0: collect title → duplicate check ──────────────────
                case 0:
                    state.SetValue(TitleKey, input);

                    var aadId = ctx.Activity.From?.AadObjectId ?? _fallbackAadId;
                    var userName = ctx.Activity.From?.Name ?? "Unknown";

                    // ── Graph enrichment: fetch reporter details silently ──────
                    var graphUser = await _graph.GetUserAsync(aadId ?? string.Empty, ct);
                    if (graphUser != null)
                    {
                        if (!string.IsNullOrEmpty(graphUser.Email))
                            state.SetValue(ReporterEmailKey, graphUser.Email);
                        if (!string.IsNullOrEmpty(graphUser.Department))
                            state.SetValue(DepartmentKey, graphUser.Department);
                    }

                    var userTickets = !string.IsNullOrEmpty(aadId)
                                        ? await _tickets.GetTicketsByAadIdAsync(aadId, ct)
                                        : Enumerable.Empty<IncidentTicket>().ToList();

                    if (!userTickets.Any())
                        userTickets = await _tickets.GetTicketsByUserAsync(userName, ct);
                    var dupResult = await _openAI.CheckDuplicateAsync(input, userTickets, ct);

                    if (dupResult.Found)
                    {
                        state.SetValue(DupIdKey, dupResult.DuplicateId!);
                        state.SetValue(DupTitleKey, dupResult.DuplicateTitle!);
                        state.SetValue(StepKey, "1");

                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"I found a similar open ticket you raised:\n\n" +
                                $"**{dupResult.DuplicateId}** — {dupResult.DuplicateTitle}\n\n" +
                                $"Is this the same issue? Reply **Yes** to track it or **No** to log a new ticket.",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    else
                    {
                        state.SetValue(StepKey, "2");
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "Can you provide more details? *(e.g. when it started, any error messages, steps to reproduce)*",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    break;

                // ── Step 1: duplicate confirmation ───────────────────────────
                case 1:
                    if (input.StartsWith("y", StringComparison.OrdinalIgnoreCase))
                    {
                        var dupId = state.GetValue<string>(DupIdKey) ?? string.Empty;
                        ClearState(state);

                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"Got it — I've noted you're also affected by **{dupId}**. " +
                                $"The assigned technician will be in touch.",
                                inputHint: InputHints.IgnoringInput),
                            cancellationToken: ct);

                        await ctx.SendActivityAsync(
                            Activity.CreateEndOfConversationActivity(),
                            cancellationToken: ct);
                    }
                    else if (input.StartsWith("n", StringComparison.OrdinalIgnoreCase))
                    {
                        state.SetValue(StepKey, "2");

                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "Can you provide more details? *(e.g. when it started, any error messages, steps to reproduce)*",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    else
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "Please reply **Yes** to track the existing ticket or **No** to log a new one.",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    break;

                // ── Step 2: collect description ──────────────────────────────
                case 2:
                    state.SetValue(DescriptionKey, input);
                    state.SetValue(StepKey, "3");

                    await ctx.SendActivityAsync(
                        MessageFactory.Text(
                            "What is the **priority**?\n\nPlease reply with: **Low**, **Medium**, **High**, or **Critical**.",
                            inputHint: InputHints.ExpectingInput),
                        cancellationToken: ct);
                    break;

                // ── Step 3: collect priority → branch on Critical ────────────
                case 3:
                    if (!TicketPriority.IsValid(input))
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"**\"{input}\"** isn't a recognised priority. Please reply with: **Low**, **Medium**, **High**, or **Critical**.",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                        return;
                    }

                    var normPriority = TicketPriority.Normalise(input);
                    state.SetValue(PriorityKey, normPriority);

                    if (normPriority == TicketPriority.Critical)
                    {
                        // ── Adaptive branch: extra escalation questions ───────
                        state.SetValue(StepKey, "4");
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "⚠️ **Critical priority selected.**\n\n" +
                                "How many users are affected?\n\n" +
                                "Reply with: **Just me**, **My team**, or **Entire organisation**",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    else
                    {
                        // ── Standard path: skip escalation steps ─────────────
                        state.SetValue(StepKey, "6");
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "Which system is affected? *(e.g. Teams, SharePoint, Outlook, VPN, Laptop)*",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    break;

                // ── Step 4: [Critical only] users affected ───────────────────
                case 4:
                    var usersAffectedOptions = new[] { "just me", "my team", "entire organisation", "entire organization" };
                    if (!usersAffectedOptions.Any(o => input.Contains(o, StringComparison.OrdinalIgnoreCase)))
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"Please reply with: **Just me**, **My team**, or **Entire organisation**.",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                        return;
                    }

                    // Normalise to a clean label
                    var usersAffectedLabel = input.Contains("organisation", StringComparison.OrdinalIgnoreCase)
                                         || input.Contains("organization", StringComparison.OrdinalIgnoreCase)
                                            ? "Entire organisation"
                                            : input.Contains("team", StringComparison.OrdinalIgnoreCase)
                                            ? "My team"
                                            : "Just me";

                    state.SetValue(UsersAffectedKey, usersAffectedLabel);
                    state.SetValue(StepKey, "5");

                    await ctx.SendActivityAsync(
                        MessageFactory.Text(
                            "Is there a workaround available right now?\n\nReply **Yes** or **No**.",
                            inputHint: InputHints.ExpectingInput),
                        cancellationToken: ct);
                    break;

                // ── Step 5: [Critical only] workaround available ─────────────
                case 5:
                    if (!input.StartsWith("y", StringComparison.OrdinalIgnoreCase) &&
                        !input.StartsWith("n", StringComparison.OrdinalIgnoreCase))
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "Please reply **Yes** or **No**.",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                        return;
                    }

                    var hasWorkaround = input.StartsWith("y", StringComparison.OrdinalIgnoreCase);
                    state.SetValue(WorkaroundKey, hasWorkaround.ToString());
                    state.SetValue(StepKey, "6");

                    await ctx.SendActivityAsync(
                        MessageFactory.Text(
                            "Which system is affected? *(e.g. Teams, SharePoint, Outlook, VPN, Laptop)*",
                            inputHint: InputHints.ExpectingInput),
                        cancellationToken: ct);
                    break;

                // ── Step 6: collect affected system → show confirmation ───────
                case 6:
                    state.SetValue(SystemKey, input);
                    state.SetValue(StepKey, "7");

                    var titleVal = state.GetValue<string>(TitleKey) ?? string.Empty;
                    var descVal = state.GetValue<string>(DescriptionKey) ?? string.Empty;
                    var priorityVal = state.GetValue<string>(PriorityKey) ?? string.Empty;
                    var usersVal = state.GetValue<string>(UsersAffectedKey);
                    var workaroundRaw = state.GetValue<string>(WorkaroundKey);
                    var workaroundVal = workaroundRaw is not null
                                        ? (bool.Parse(workaroundRaw) ? "Yes" : "No")
                                        : null;
                    var isEscalated = priorityVal == TicketPriority.Critical
                                     && workaroundRaw is not null
                                     && !bool.Parse(workaroundRaw);
                    var emailVal = state.GetValue<string>(ReporterEmailKey);
                    var departmentVal = state.GetValue<string>(DepartmentKey);

                    // Build confirmation table — add escalation/graph rows only when present
                    var confirmRows = new System.Text.StringBuilder();
                    confirmRows.AppendLine($"| **Issue** | {titleVal} |");
                    confirmRows.AppendLine($"| **Details** | {descVal} |");
                    confirmRows.AppendLine($"| **Priority** | {priorityVal} |");
                    if (!string.IsNullOrEmpty(emailVal))
                        confirmRows.AppendLine($"| **Reporter Email** | {emailVal} |");
                    if (!string.IsNullOrEmpty(departmentVal))
                        confirmRows.AppendLine($"| **Department** | {departmentVal} |");
                    if (usersVal is not null)
                        confirmRows.AppendLine($"| **Users Affected** | {usersVal} |");
                    if (workaroundVal is not null)
                        confirmRows.AppendLine($"| **Workaround** | {workaroundVal} |");
                    confirmRows.AppendLine($"| **Affected System** | {input} |");
                    if (isEscalated)
                        confirmRows.AppendLine($"| **Escalation** | 🚨 Immediate escalation required |");

                    await ctx.SendActivityAsync(
                        MessageFactory.Text(
                            $"Here's what I have:\n\n" +
                            $"| Field | Value |\n" +
                            $"|---|---|\n" +
                            confirmRows.ToString() +
                            $"\nShall I log this ticket? Reply **Yes** to confirm or **No** to cancel.",
                            inputHint: InputHints.ExpectingInput),
                        cancellationToken: ct);
                    break;

                // ── Step 7: confirmation → create ticket or cancel ───────────
                case 7:
                    if (input.StartsWith("y", StringComparison.OrdinalIgnoreCase))
                    {
                        var title = state.GetValue<string>(TitleKey) ?? string.Empty;
                        var description = state.GetValue<string>(DescriptionKey) ?? string.Empty;
                        var priority = state.GetValue<string>(PriorityKey) ?? string.Empty;
                        var system = state.GetValue<string>(SystemKey) ?? string.Empty;
                        var usersAff = state.GetValue<string>(UsersAffectedKey);
                        var wkRaw = state.GetValue<string>(WorkaroundKey);
                        var workaround = wkRaw is not null ? bool.Parse(wkRaw) : (bool?)null;
                        var escalate = priority == TicketPriority.Critical && workaround == false;
                        var email = state.GetValue<string>(ReporterEmailKey);
                        var department = state.GetValue<string>(DepartmentKey);

                        ClearState(state);

                        var ticket = await _tickets.CreateTicketAsync(new CreateTicketRequest
                        {
                            Title = title,
                            Description = description,
                            Priority = priority,
                            AffectedSystem = system,
                            ReportedBy = ctx.Activity.From?.Name ?? "Unknown",
                            ReporterAadId = ctx.Activity.From?.AadObjectId ?? _fallbackAadId,
                            ReporterEmail = email,
                            Department = department,
                            UsersAffected = usersAff,
                            WorkaroundAvailable = workaround,
                            IsEscalated = escalate
                        }, ct);

                        // SLA: Critical with no workaround → 15 min, with workaround → 30 min
                        var sla = (ticket.Priority, ticket.IsEscalated) switch
                        {
                            (TicketPriority.Critical, true) => "**15 minutes** 🚨",
                            (TicketPriority.Critical, false) => "**30 minutes**",
                            (TicketPriority.High, _) => "**1 hour**",
                            (TicketPriority.Medium, _) => "**4 hours**",
                            _ => "**next business day**"
                        };

                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                $"Your incident has been logged.\n\n" +
                                $"| Field | Value |\n" +
                                $"|---|---|\n" +
                                $"| **Ticket ID** | {ticket.Id} |\n" +
                                $"| **Issue** | {ticket.Title} |\n" +
                                $"| **Priority** | {ticket.Priority} |\n" +
                                $"| **Affected System** | {ticket.AffectedSystem} |\n" +
                                $"| **Status** | {ticket.Status} |\n\n" +
                                $"A technician will contact you within {sla}.",
                                inputHint: InputHints.IgnoringInput),
                            cancellationToken: ct);

                        await ctx.SendActivityAsync(
                            Activity.CreateEndOfConversationActivity(),
                            cancellationToken: ct);
                    }
                    else if (input.StartsWith("n", StringComparison.OrdinalIgnoreCase))
                    {
                        ClearState(state);

                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "No problem — your incident report has been cancelled. Come back any time if you need to log an issue.",
                                inputHint: InputHints.IgnoringInput),
                            cancellationToken: ct);

                        await ctx.SendActivityAsync(
                            Activity.CreateEndOfConversationActivity(),
                            cancellationToken: ct);
                    }
                    else
                    {
                        await ctx.SendActivityAsync(
                            MessageFactory.Text(
                                "Please reply **Yes** to confirm or **No** to cancel.",
                                inputHint: InputHints.ExpectingInput),
                            cancellationToken: ct);
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[SKILL] *** EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            throw;
        }
    }

    // ── Check ticket status ──────────────────────────────────────────────────

    private async Task OnCheckTicketStatusAsync(ITurnContext ctx, ITurnState state, CancellationToken ct)
    {
        // Copilot Studio passes ticketId via Activity.Value when manifest has a value schema.
        // Fall back to Activity.Text for direct callers.
        var ticketId = string.Empty;
        if (ctx.Activity.Value is System.Text.Json.JsonElement json &&
            json.TryGetProperty("ticketId", out var prop))
        {
            ticketId = prop.GetString()?.Trim() ?? string.Empty;
        }
        if (string.IsNullOrEmpty(ticketId))
            ticketId = ctx.Activity.Text?.Trim() ?? string.Empty;

        Console.WriteLine($"[SKILL][CHECK] ticketId='{ticketId}'");

        if (string.IsNullOrEmpty(ticketId))
        {
            await ctx.SendActivityAsync(
                MessageFactory.Text(
                    "Please provide a ticket ID (e.g. **INC-001**).",
                    inputHint: InputHints.IgnoringInput),
                cancellationToken: ct);

            await ctx.SendActivityAsync(Activity.CreateEndOfConversationActivity(), cancellationToken: ct);
            return;
        }

        var ticket = await _tickets.GetTicketByIdAsync(ticketId, ct);

        if (ticket is null)
        {
            await ctx.SendActivityAsync(
                MessageFactory.Text(
                    $"No ticket found with ID **{ticketId}**. Please check the ID and try again.",
                    inputHint: InputHints.IgnoringInput),
                cancellationToken: ct);
        }
        else
        {
            var assignedLine = !string.IsNullOrEmpty(ticket.AssignedTo)
                ? $"| **Assigned To** | {ticket.AssignedTo} |\n"
                : string.Empty;

            var resolutionLine = !string.IsNullOrEmpty(ticket.Resolution)
                ? $"| **Resolution** | {ticket.Resolution} |\n"
                : string.Empty;

            await ctx.SendActivityAsync(
                MessageFactory.Text(
                    $"Here is the current status for **{ticket.Id}**:\n\n" +
                    $"| Field | Value |\n" +
                    $"|---|---|\n" +
                    $"| **Ticket ID** | {ticket.Id} |\n" +
                    $"| **Issue** | {ticket.Title} |\n" +
                    $"| **Priority** | {ticket.Priority} |\n" +
                    $"| **Affected System** | {ticket.AffectedSystem} |\n" +
                    $"| **Status** | {ticket.Status} |\n" +
                    $"| **Reported By** | {ticket.ReportedBy} |\n" +
                    $"| **Created** | {ticket.CreatedAt:dd MMM yyyy HH:mm} UTC |\n" +
                    assignedLine +
                    resolutionLine,
                    inputHint: InputHints.IgnoringInput),
                cancellationToken: ct);
        }

        await ctx.SendActivityAsync(Activity.CreateEndOfConversationActivity(), cancellationToken: ct);
    }

    // ── End of conversation (Copilot Studio cancels mid-flow) ────────────────

    private Task OnEndOfConversationAsync(ITurnContext ctx, ITurnState state, CancellationToken ct)
    {
        ClearState(state);
        return Task.CompletedTask;
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static int GetInt(ITurnState state, string key)
    {
        var raw = state.GetValue<string>(key);
        return int.TryParse(raw, out var n) ? n : 0;
    }

    private static void ClearState(ITurnState state)
    {
        state.DeleteValue(StepKey);
        state.DeleteValue(TitleKey);
        state.DeleteValue(DescriptionKey);
        state.DeleteValue(PriorityKey);
        state.DeleteValue(SystemKey);
        state.DeleteValue(DupIdKey);
        state.DeleteValue(DupTitleKey);
        state.DeleteValue(UsersAffectedKey);
        state.DeleteValue(WorkaroundKey);
        state.DeleteValue(ReporterEmailKey);
        state.DeleteValue(DepartmentKey);
    }
}
