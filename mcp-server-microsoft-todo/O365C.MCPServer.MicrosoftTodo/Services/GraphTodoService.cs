using Microsoft.Graph;
using Microsoft.Graph.Models;
using O365C.MCPServer.MicrosoftTodo.Models;
using System.Security.Claims;
using System.Text;

// Aliases to avoid name clashes between our models and Graph SDK models
using AppTodoList = O365C.MCPServer.MicrosoftTodo.Models.TodoList;
using AppTodoTask = O365C.MCPServer.MicrosoftTodo.Models.TodoTask;
using AppChecklistItem = O365C.MCPServer.MicrosoftTodo.Models.TodoChecklistItem;
using AppLinkedResource = O365C.MCPServer.MicrosoftTodo.Models.TodoLinkedResource;
using AppTaskAttachment = O365C.MCPServer.MicrosoftTodo.Models.TodoTaskAttachment;
using GraphTodoTask = Microsoft.Graph.Models.TodoTask;
using GraphTodoTaskList = Microsoft.Graph.Models.TodoTaskList;
using GraphChecklistItem = Microsoft.Graph.Models.ChecklistItem;
using GraphLinkedResource = Microsoft.Graph.Models.LinkedResource;
using GraphTaskFileAttachment = Microsoft.Graph.Models.TaskFileAttachment;
using GraphAttachment = Microsoft.Graph.Models.AttachmentBase;
using GraphTaskStatus = Microsoft.Graph.Models.TaskStatus;
using GraphImportance = Microsoft.Graph.Models.Importance;
using GraphDateTimeTimeZone = Microsoft.Graph.Models.DateTimeTimeZone;

namespace O365C.MCPServer.MicrosoftTodo.Services;

/// <summary>
/// Calls Microsoft Graph on behalf of the signed-in user using the OBO flow.
/// GraphServiceClient is injected by Microsoft.Identity.Web and transparently
/// performs the On-Behalf-Of token exchange for each request.
/// </summary>
public class GraphTodoService(
    GraphServiceClient graph,
    ILogger<GraphTodoService> logger,
    IHttpContextAccessor http,
    IHttpClientFactory httpClientFactory)
{
    // -------------------------------------------------------------------------
    // Cross-cutting helpers
    // -------------------------------------------------------------------------

    private void LogCallerIdentity(string operation)
    {
        var claims = http.HttpContext?.User.Claims;
        var name = claims?.FirstOrDefault(c => c.Type == "name")?.Value ?? "unknown";
        var upn = claims?.FirstOrDefault(c => c.Type == "preferred_username")?.Value
                ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Upn)?.Value
                ?? "unknown";
        logger.LogInformation("[OBO] {Operation} → acting on behalf of: {Name} ({Upn})",
            operation, name, upn);
    }

    // -------------------------------------------------------------------------
    // To Do list operations (list CRUD)
    // -------------------------------------------------------------------------

    public async Task<List<AppTodoList>> GetTodoListsAsync()
    {
        LogCallerIdentity("GetTodoLists");
        var response = await graph.Me.Todo.Lists.GetAsync();
        return response?.Value?.Select(MapTodoList).ToList() ?? [];
    }

    public async Task<AppTodoList?> GetTodoListByIdAsync(string listId)
    {
        LogCallerIdentity("GetTodoListById");
        var list = await graph.Me.Todo.Lists[listId].GetAsync();
        return list is null ? null : MapTodoList(list);
    }

    public async Task<AppTodoList> CreateTodoListAsync(string displayName)
    {
        LogCallerIdentity("CreateTodoList");
        var created = await graph.Me.Todo.Lists.PostAsync(new GraphTodoTaskList
        {
            DisplayName = displayName
        });
        return MapTodoList(created!);
    }

    public async Task<AppTodoList> UpdateTodoListAsync(string listId, string displayName)
    {
        LogCallerIdentity("UpdateTodoList");
        var updated = await graph.Me.Todo.Lists[listId].PatchAsync(new GraphTodoTaskList
        {
            DisplayName = displayName
        });
        return MapTodoList(updated!);
    }

    public async Task DeleteTodoListAsync(string listId)
    {
        LogCallerIdentity("DeleteTodoList");
        await graph.Me.Todo.Lists[listId].DeleteAsync();
    }

    // -------------------------------------------------------------------------
    // To Do task operations (task CRUD + queries)
    // -------------------------------------------------------------------------

    public async Task<List<AppTodoTask>> GetTasksByListAsync(string listId)
    {
        LogCallerIdentity("GetTasksByList");
        var response = await graph.Me.Todo.Lists[listId].Tasks.GetAsync();
        return response?.Value?
            .Select(t => MapTodoTask(t, listId))
            .ToList() ?? [];
    }

    public async Task<AppTodoTask?> GetTaskByIdAsync(string listId, string taskId)
    {
        LogCallerIdentity("GetTaskById");
        var task = await graph.Me.Todo.Lists[listId].Tasks[taskId].GetAsync();
        return task is null ? null : MapTodoTask(task, listId);
    }

    public async Task<List<AppTodoTask>> GetAllTodoTasksAsync()
    {
        LogCallerIdentity("GetAllTodoTasks");
        var lists = await GetTodoListsAsync();
        var results = await Task.WhenAll(lists.Select(l => GetTasksByListAsync(l.Id)));
        return [.. results.SelectMany(t => t)];
    }

    public async Task<AppTodoTask> CreateTaskAsync(string listId, string title, string? body = null)
    {
        LogCallerIdentity("CreateTask");
        var newTask = new GraphTodoTask
        {
            Title = title,
            Body = body != null ? new ItemBody { Content = body, ContentType = BodyType.Text } : null
        };
        var created = await graph.Me.Todo.Lists[listId].Tasks.PostAsync(newTask);
        return MapTodoTask(created!, listId);
    }

    public async Task<AppTodoTask> CompleteTaskAsync(string listId, string taskId)
    {
        LogCallerIdentity("CompleteTask");
        var patch = new GraphTodoTask { Status = GraphTaskStatus.Completed };
        var updated = await graph.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(patch);
        return MapTodoTask(updated!, listId);
    }

    public async Task<AppTodoTask> UpdateTaskAsync(
        string listId,
        string taskId,
        string? title = null,
        string? body = null,
        string? status = null,
        string? importance = null,
        DateTime? dueDateTime = null,
        bool clearDueDate = false)
    {
        LogCallerIdentity("UpdateTask");

        var patch = new GraphTodoTask();
        if (!string.IsNullOrWhiteSpace(title))
            patch.Title = title;

        if (body is not null)
            patch.Body = new ItemBody { Content = body, ContentType = BodyType.Text };

        if (!string.IsNullOrWhiteSpace(status))
            patch.Status = ParseTaskStatus(status);

        if (!string.IsNullOrWhiteSpace(importance))
            patch.Importance = ParseImportance(importance);

        if (clearDueDate)
            patch.DueDateTime = null;
        else if (dueDateTime.HasValue)
            patch.DueDateTime = new GraphDateTimeTimeZone
            {
                DateTime = dueDateTime.Value.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC"
            };

        var updated = await graph.Me.Todo.Lists[listId].Tasks[taskId].PatchAsync(patch);
        return MapTodoTask(updated!, listId);
    }

    public async Task DeleteTaskAsync(string listId, string taskId)
    {
        LogCallerIdentity("DeleteTask");
        await graph.Me.Todo.Lists[listId].Tasks[taskId].DeleteAsync();
    }

    // -------------------------------------------------------------------------
    // Checklist item operations
    // -------------------------------------------------------------------------

    public async Task<List<AppChecklistItem>> GetChecklistItemsAsync(string listId, string taskId)
    {
        LogCallerIdentity("GetChecklistItems");
        var response = await graph.Me.Todo.Lists[listId].Tasks[taskId].ChecklistItems.GetAsync();
        return response?.Value?.Select(MapChecklistItem).ToList() ?? [];
    }

    public async Task<AppChecklistItem> CreateChecklistItemAsync(
        string listId,
        string taskId,
        string displayName,
        bool isChecked = false)
    {
        LogCallerIdentity("CreateChecklistItem");
        var created = await graph.Me.Todo.Lists[listId].Tasks[taskId].ChecklistItems.PostAsync(new GraphChecklistItem
        {
            DisplayName = displayName,
            IsChecked = isChecked
        });
        return MapChecklistItem(created!);
    }

    public async Task<AppChecklistItem> UpdateChecklistItemAsync(
        string listId,
        string taskId,
        string checklistItemId,
        string? displayName = null,
        bool? isChecked = null)
    {
        LogCallerIdentity("UpdateChecklistItem");
        var patch = new GraphChecklistItem();
        if (!string.IsNullOrWhiteSpace(displayName))
            patch.DisplayName = displayName;
        if (isChecked.HasValue)
            patch.IsChecked = isChecked.Value;

        var updated = await graph.Me.Todo.Lists[listId].Tasks[taskId].ChecklistItems[checklistItemId].PatchAsync(patch);
        return MapChecklistItem(updated!);
    }

    public async Task DeleteChecklistItemAsync(string listId, string taskId, string checklistItemId)
    {
        LogCallerIdentity("DeleteChecklistItem");
        await graph.Me.Todo.Lists[listId].Tasks[taskId].ChecklistItems[checklistItemId].DeleteAsync();
    }

    // -------------------------------------------------------------------------
    // Linked resource operations
    // -------------------------------------------------------------------------

    public async Task<List<AppLinkedResource>> GetLinkedResourcesAsync(string listId, string taskId)
    {
        LogCallerIdentity("GetLinkedResources");
        var response = await graph.Me.Todo.Lists[listId].Tasks[taskId].LinkedResources.GetAsync();
        return response?.Value?.Select(MapLinkedResource).ToList() ?? [];
    }

    public async Task<AppLinkedResource> CreateLinkedResourceAsync(
        string listId,
        string taskId,
        string displayName,
        string webUrl,
        string? applicationName = null,
        string? externalId = null)
    {
        LogCallerIdentity("CreateLinkedResource");
        var created = await graph.Me.Todo.Lists[listId].Tasks[taskId].LinkedResources.PostAsync(new GraphLinkedResource
        {
            DisplayName = displayName,
            WebUrl = webUrl,
            ApplicationName = string.IsNullOrWhiteSpace(applicationName) ? "Copilot Studio" : applicationName,
            ExternalId = externalId
        });
        return MapLinkedResource(created!);
    }

    public async Task DeleteLinkedResourceAsync(string listId, string taskId, string linkedResourceId)
    {
        LogCallerIdentity("DeleteLinkedResource");
        await graph.Me.Todo.Lists[listId].Tasks[taskId].LinkedResources[linkedResourceId].DeleteAsync();
    }

    // -------------------------------------------------------------------------
    // Attachment operations
    // -------------------------------------------------------------------------

    public async Task<List<AppTaskAttachment>> GetTaskAttachmentsAsync(string listId, string taskId)
    {
        LogCallerIdentity("GetTaskAttachments");
        var response = await graph.Me.Todo.Lists[listId].Tasks[taskId].Attachments.GetAsync();
        return response?.Value?.Select(MapAttachment).ToList() ?? [];
    }

    public async Task<AppTaskAttachment> AddTaskFileAttachmentAsync(
        string listId,
        string taskId,
        string name,
        byte[] contentBytes,
        string contentType = "application/octet-stream")
    {
        LogCallerIdentity("AddTaskFileAttachment");
        var attachment = new GraphTaskFileAttachment
        {
            Name = name,
            ContentType = contentType,
            ContentBytes = contentBytes
        };
        var created = await graph.Me.Todo.Lists[listId].Tasks[taskId].Attachments.PostAsync(attachment);
        return MapAttachment(created!);
    }

    public async Task<AppTaskAttachment> AddTaskFileAttachmentFromUrlAsync(
        string listId,
        string taskId,
        string fileUrl,
        string? name = null,
        string? contentType = null)
    {
        LogCallerIdentity("AddTaskFileAttachmentFromUrl");

        if (!Uri.TryCreate(fileUrl, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            throw new ArgumentException("fileUrl must be an absolute http/https URL.", nameof(fileUrl));
        }

        // For private SharePoint/OneDrive links, use Graph so delegated OBO identity is honored.
        if (IsSharePointOrOneDriveHost(uri.Host))
        {
            try
            {
                return await AddTaskFileAttachmentFromGraphShareUrlAsync(listId, taskId, uri, name, contentType);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex,
                    "Graph-based file download failed for {Host}; attempting HTTP fallback.",
                    uri.Host);
            }
        }

        return await AddTaskFileAttachmentFromHttpUrlAsync(listId, taskId, uri, name, contentType);
    }

    private async Task<AppTaskAttachment> AddTaskFileAttachmentFromGraphShareUrlAsync(
        string listId,
        string taskId,
        Uri uri,
        string? name,
        string? contentType)
    {
        var shareId = EncodeGraphSharingUrl(uri.AbsoluteUri);
        var driveItem = await graph.Shares[shareId].DriveItem.GetAsync();
        var contentStream = await graph.Shares[shareId].DriveItem.Content.GetAsync();

        if (contentStream is null)
            throw new InvalidOperationException("Graph returned empty content stream for the requested file URL.");

        using var memory = new MemoryStream();
        await contentStream.CopyToAsync(memory);
        var bytes = memory.ToArray();
        if (bytes.Length == 0)
            throw new InvalidOperationException("Downloaded file is empty.");

        var finalName = string.IsNullOrWhiteSpace(name)
            ? driveItem?.Name ?? InferFileNameFromUrl(uri)
            : name!;

        var finalContentType = string.IsNullOrWhiteSpace(contentType)
            ? driveItem?.File?.MimeType ?? "application/octet-stream"
            : contentType;

        return await AddTaskFileAttachmentAsync(listId, taskId, finalName, bytes, finalContentType);
    }

    private async Task<AppTaskAttachment> AddTaskFileAttachmentFromHttpUrlAsync(
        string listId,
        string taskId,
        Uri uri,
        string? name,
        string? contentType)
    {
        using var client = httpClientFactory.CreateClient();
        using var response = await client.GetAsync(uri);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Failed to download attachment from URL. HTTP {(int)response.StatusCode} {response.ReasonPhrase}.");
        }

        var bytes = await response.Content.ReadAsByteArrayAsync();
        if (bytes.Length == 0)
            throw new InvalidOperationException("Downloaded file is empty.");

        var finalName = string.IsNullOrWhiteSpace(name)
            ? InferFileNameFromUrl(uri)
            : name!;

        var finalContentType = string.IsNullOrWhiteSpace(contentType)
            ? response.Content.Headers.ContentType?.MediaType ?? "application/octet-stream"
            : contentType;

        return await AddTaskFileAttachmentAsync(listId, taskId, finalName, bytes, finalContentType);
    }

    public async Task DeleteTaskAttachmentAsync(string listId, string taskId, string attachmentId)
    {
        LogCallerIdentity("DeleteTaskAttachment");
        await graph.Me.Todo.Lists[listId].Tasks[taskId].Attachments[attachmentId].DeleteAsync();
    }

    // -------------------------------------------------------------------------
    // Mapping + parsing helpers
    // -------------------------------------------------------------------------

    private static AppTodoList MapTodoList(GraphTodoTaskList l) => new()
    {
        Id = l.Id ?? "",
        DisplayName = l.DisplayName ?? "",
        IsOwner = l.IsOwner ?? false,
        IsShared = l.IsShared ?? false
    };

    private static AppTodoTask MapTodoTask(GraphTodoTask t, string listId) => new()
    {
        Id = t.Id ?? "",
        ListId = listId,
        Title = t.Title ?? "",
        Status = t.Status switch
        {
            GraphTaskStatus.NotStarted => "notStarted",
            GraphTaskStatus.InProgress => "inProgress",
            GraphTaskStatus.Completed => "completed",
            _ => "notStarted"
        },
        Importance = t.Importance switch
        {
            GraphImportance.Low => "low",
            GraphImportance.Normal => "normal",
            GraphImportance.High => "high",
            _ => "normal"
        },
        DueDateTime = ParseDateTimeTimeZone(t.DueDateTime),
        CompletedDateTime = ParseDateTimeTimeZone(t.CompletedDateTime),
        CreatedDateTime = t.CreatedDateTime?.UtcDateTime ?? DateTime.UtcNow,
        Body = t.Body?.Content
    };

    private static AppChecklistItem MapChecklistItem(GraphChecklistItem item) => new()
    {
        Id = item.Id ?? "",
        DisplayName = item.DisplayName ?? "",
        IsChecked = item.IsChecked ?? false,
        CreatedDateTime = item.CreatedDateTime?.UtcDateTime,
        CheckedDateTime = item.CheckedDateTime?.UtcDateTime
    };

    private static AppLinkedResource MapLinkedResource(GraphLinkedResource resource) => new()
    {
        Id = resource.Id ?? "",
        DisplayName = resource.DisplayName ?? "",
        ApplicationName = resource.ApplicationName,
        ExternalId = resource.ExternalId,
        WebUrl = resource.WebUrl
    };

    private static AppTaskAttachment MapAttachment(GraphAttachment attachment) => new()
    {
        Id = attachment.Id ?? "",
        Name = attachment.Name ?? "",
        ContentType = attachment.ContentType,
        Size = attachment.Size,
        LastModifiedDateTime = attachment.LastModifiedDateTime?.UtcDateTime,
        AttachmentType = attachment.OdataType
    };

    private static DateTime? ParseDateTimeTimeZone(GraphDateTimeTimeZone? dtz)
    {
        if (dtz?.DateTime is null) return null;
        return DateTime.TryParse(dtz.DateTime, out var dt) ? dt : null;
    }

    private static string InferFileNameFromUrl(Uri uri)
    {
        var lastSegment = uri.Segments.LastOrDefault();
        if (string.IsNullOrWhiteSpace(lastSegment))
            return "downloaded-file";

        var candidate = Uri.UnescapeDataString(lastSegment.Trim('/'));
        return string.IsNullOrWhiteSpace(candidate) ? "downloaded-file" : candidate;
    }

    private static bool IsSharePointOrOneDriveHost(string host)
    {
        var normalized = host.ToLowerInvariant();
        return normalized.EndsWith(".sharepoint.com")
            || normalized.EndsWith("-my.sharepoint.com")
            || normalized.Contains("sharepoint.com");
    }

    private static string EncodeGraphSharingUrl(string absoluteUrl)
    {
        var bytes = Encoding.UTF8.GetBytes(absoluteUrl);
        var base64 = Convert.ToBase64String(bytes)
            .TrimEnd('=')
            .Replace('/', '_')
            .Replace('+', '-');
        return $"u!{base64}";
    }

    private static GraphTaskStatus ParseTaskStatus(string status) => status.Trim().ToLowerInvariant() switch
    {
        "notstarted" => GraphTaskStatus.NotStarted,
        "not_started" => GraphTaskStatus.NotStarted,
        "inprogress" => GraphTaskStatus.InProgress,
        "in_progress" => GraphTaskStatus.InProgress,
        "completed" => GraphTaskStatus.Completed,
        "waitingonothers" => GraphTaskStatus.WaitingOnOthers,
        "waiting_on_others" => GraphTaskStatus.WaitingOnOthers,
        "deferred" => GraphTaskStatus.Deferred,
        _ => GraphTaskStatus.NotStarted
    };

    private static GraphImportance ParseImportance(string importance) => importance.Trim().ToLowerInvariant() switch
    {
        "low" => GraphImportance.Low,
        "normal" => GraphImportance.Normal,
        "high" => GraphImportance.High,
        _ => GraphImportance.Normal
    };
}