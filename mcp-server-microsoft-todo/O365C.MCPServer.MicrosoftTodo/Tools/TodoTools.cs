using ModelContextProtocol.Server;
using O365C.MCPServer.MicrosoftTodo.Services;
using System.ComponentModel;
using System.Security.Claims;
using System.Text;
using System.Text.Json;

namespace O365C.MCPServer.MicrosoftTodo.Tools;

[McpServerToolType]
public class TodoTools(GraphTodoService graphService)
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    // ---------------------------------------------------------------------
    // To Do list tools
    // ---------------------------------------------------------------------

    [McpServerTool(Name = "todo_get_my_lists"), Description("Returns all Microsoft To Do lists for the signed-in user.")]
    public async Task<string> GetMyTodoLists()
    {
        var lists = await graphService.GetTodoListsAsync();
        return JsonSerializer.Serialize(lists, JsonOptions);
    }

    [McpServerTool(Name = "todo_get_list_by_id"), Description("Returns one Microsoft To Do list by listId.")]
    public async Task<string> GetTodoListById(
        [Description("Microsoft Graph To Do list ID. Use todo_get_my_lists first.")] string listId)
    {
        var list = await graphService.GetTodoListByIdAsync(listId);
        return list is null
            ? $"List '{listId}' was not found."
            : JsonSerializer.Serialize(list, JsonOptions);
    }

    [McpServerTool(Name = "todo_create_list"), Description("Creates a new Microsoft To Do list.")]
    public async Task<string> CreateTodoList(
        [Description("Display name for the new To Do list.")] string displayName)
    {
        var list = await graphService.CreateTodoListAsync(displayName);
        return $"To Do list created: '{list.DisplayName}' (Id: {list.Id})";
    }

    [McpServerTool(Name = "todo_update_list"), Description("Updates an existing Microsoft To Do list display name.")]
    public async Task<string> UpdateTodoList(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("New display name for the list.")] string displayName)
    {
        var list = await graphService.UpdateTodoListAsync(listId, displayName);
        return $"To Do list updated: '{list.DisplayName}' (Id: {list.Id})";
    }

    [McpServerTool(Name = "todo_delete_list"), Description("Deletes a Microsoft To Do list by listId.")]
    public async Task<string> DeleteTodoList(
        [Description("Microsoft Graph To Do list ID.")] string listId)
    {
        await graphService.DeleteTodoListAsync(listId);
        return $"To Do list '{listId}' deleted.";
    }

    // ---------------------------------------------------------------------
    // To Do task tools
    // ---------------------------------------------------------------------

    [McpServerTool(Name = "todo_get_tasks_by_list"), Description("Returns all tasks in a specific Microsoft To Do list.")]
    public async Task<string> GetTasksByList(
        [Description("Microsoft Graph To Do list ID.")] string listId)
    {
        var tasks = await graphService.GetTasksByListAsync(listId);
        if (tasks.Count == 0)
            return $"No tasks found in list '{listId}'.";
        return JsonSerializer.Serialize(tasks, JsonOptions);
    }

    [McpServerTool(Name = "todo_get_task_by_id"), Description("Returns one task from a specific To Do list.")]
    public async Task<string> GetTaskById(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId)
    {
        var task = await graphService.GetTaskByIdAsync(listId, taskId);
        return task is null
            ? $"Task '{taskId}' was not found in list '{listId}'."
            : JsonSerializer.Serialize(task, JsonOptions);
    }

    [McpServerTool(Name = "todo_create_task"), Description("Creates a new task in a Microsoft To Do list.")]
    public async Task<string> CreateTask(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Task title.")] string title,
        [Description("Optional plain-text task body.")] string? body = null)
    {
        var task = await graphService.CreateTaskAsync(listId, title, body);
        return $"Task created: '{task.Title}' (Id: {task.Id})";
    }

    [McpServerTool(Name = "todo_update_task"), Description("Updates a task. Optional fields: title, body, status, importance, dueDateTimeUtc, and clearDueDate.")]
    public async Task<string> UpdateTask(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Optional new task title.")] string? title = null,
        [Description("Optional new plain-text task body.")] string? body = null,
        [Description("Optional status: notStarted|inProgress|completed|waitingOnOthers|deferred.")] string? status = null,
        [Description("Optional importance: low|normal|high.")] string? importance = null,
        [Description("Optional UTC due date/time in ISO 8601 format, for example 2026-06-14T16:00:00Z.")] string? dueDateTimeUtc = null,
        [Description("Set true to remove due date.")]
        bool clearDueDate = false)
    {
        DateTime? parsedDue = null;
        if (!string.IsNullOrWhiteSpace(dueDateTimeUtc) && DateTime.TryParse(dueDateTimeUtc, out var dt))
            parsedDue = dt.ToUniversalTime();

        var task = await graphService.UpdateTaskAsync(
            listId,
            taskId,
            title,
            body,
            status,
            importance,
            parsedDue,
            clearDueDate);

        return $"Task updated: '{task.Title}' (Id: {task.Id})";
    }

    [McpServerTool(Name = "todo_delete_task"), Description("Deletes a task from a To Do list.")]
    public async Task<string> DeleteTask(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId)
    {
        await graphService.DeleteTaskAsync(listId, taskId);
        return $"Task '{taskId}' deleted from list '{listId}'.";
    }

    [McpServerTool(Name = "todo_complete_task"), Description("Marks a task as completed.")]
    public async Task<string> CompleteTask(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId)
    {
        var task = await graphService.CompleteTaskAsync(listId, taskId);
        return $"Task '{task.Title}' marked as completed.";
    }

    [McpServerTool(Name = "todo_get_overdue_tasks"), Description("Returns all incomplete tasks that are past their due date across all lists.")]
    public async Task<string> GetOverdueTasks()
    {
        var all = await graphService.GetAllTodoTasksAsync();
        var overdue = all
            .Where(t => t.Status != "completed"
                     && t.DueDateTime.HasValue
                     && t.DueDateTime.Value.Date < DateTime.Today)
            .OrderBy(t => t.DueDateTime)
            .ToList();

        return overdue.Count == 0
            ? "No overdue tasks. You are all caught up!"
            : JsonSerializer.Serialize(overdue, JsonOptions);
    }

    // ---------------------------------------------------------------------
    // Checklist item tools
    // ---------------------------------------------------------------------

    [McpServerTool(Name = "todo_get_checklist_items"), Description("Returns checklist items for a task.")]
    public async Task<string> GetChecklistItems(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId)
    {
        var items = await graphService.GetChecklistItemsAsync(listId, taskId);
        return items.Count == 0
            ? $"No checklist items found for task '{taskId}'."
            : JsonSerializer.Serialize(items, JsonOptions);
    }

    [McpServerTool(Name = "todo_create_checklist_item"), Description("Creates a checklist item in a task.")]
    public async Task<string> CreateChecklistItem(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Checklist item text.")] string displayName,
        [Description("Optional initial checked state.")] bool isChecked = false)
    {
        var item = await graphService.CreateChecklistItemAsync(listId, taskId, displayName, isChecked);
        return $"Checklist item created: '{item.DisplayName}' (Id: {item.Id})";
    }

    [McpServerTool(Name = "todo_update_checklist_item"), Description("Updates a checklist item in a task.")]
    public async Task<string> UpdateChecklistItem(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Checklist item ID.")] string checklistItemId,
        [Description("Optional updated checklist item text.")] string? displayName = null,
        [Description("Optional updated checked state.")]
        bool? isChecked = null)
    {
        var item = await graphService.UpdateChecklistItemAsync(listId, taskId, checklistItemId, displayName, isChecked);
        return $"Checklist item updated: '{item.DisplayName}' (Id: {item.Id})";
    }

    [McpServerTool(Name = "todo_delete_checklist_item"), Description("Deletes a checklist item from a task.")]
    public async Task<string> DeleteChecklistItem(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Checklist item ID.")] string checklistItemId)
    {
        await graphService.DeleteChecklistItemAsync(listId, taskId, checklistItemId);
        return $"Checklist item '{checklistItemId}' deleted from task '{taskId}'.";
    }

    // ---------------------------------------------------------------------
    // Linked resource tools
    // ---------------------------------------------------------------------

    [McpServerTool(Name = "todo_get_linked_resources"), Description("Returns linked resources for a task (web links and references).")]
    public async Task<string> GetLinkedResources(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId)
    {
        var resources = await graphService.GetLinkedResourcesAsync(listId, taskId);
        return resources.Count == 0
            ? $"No linked resources found for task '{taskId}'."
            : JsonSerializer.Serialize(resources, JsonOptions);
    }

    [McpServerTool(Name = "todo_create_linked_resource"), Description("Creates a linked resource on a task.")]
    public async Task<string> CreateLinkedResource(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Display name of the linked resource.")] string displayName,
        [Description("Target web URL.")] string webUrl,
        [Description("Optional source application name.")] string? applicationName = null,
        [Description("Optional external ID from the source system.")]
        string? externalId = null)
    {
        var resource = await graphService.CreateLinkedResourceAsync(
            listId,
            taskId,
            displayName,
            webUrl,
            applicationName,
            externalId);

        return $"Linked resource created: '{resource.DisplayName}' (Id: {resource.Id})";
    }

    [McpServerTool(Name = "todo_delete_linked_resource"), Description("Deletes a linked resource from a task.")]
    public async Task<string> DeleteLinkedResource(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Linked resource ID.")] string linkedResourceId)
    {
        await graphService.DeleteLinkedResourceAsync(listId, taskId, linkedResourceId);
        return $"Linked resource '{linkedResourceId}' deleted from task '{taskId}'.";
    }

    // ---------------------------------------------------------------------
    // Attachment tools
    // ---------------------------------------------------------------------

    [McpServerTool(Name = "todo_get_task_attachments"), Description("Returns attachments for a task.")]
    public async Task<string> GetTaskAttachments(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId)
    {
        var attachments = await graphService.GetTaskAttachmentsAsync(listId, taskId);
        return attachments.Count == 0
            ? $"No attachments found for task '{taskId}'."
            : JsonSerializer.Serialize(attachments, JsonOptions);
    }

    [McpServerTool(Name = "todo_add_task_file_attachment"), Description("Creates a file attachment on a task using base64 content.")]
    public async Task<string> AddTaskFileAttachment(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Attachment file name, including extension.")] string name,
        [Description("Raw file content encoded as base64.")] string base64Content,
        [Description("MIME type. Default is application/octet-stream.")]
        string contentType = "application/octet-stream")
    {
        var bytes = Convert.FromBase64String(base64Content);
        var attachment = await graphService.AddTaskFileAttachmentAsync(listId, taskId, name, bytes, contentType);
        return $"Attachment created: '{attachment.Name}' (Id: {attachment.Id})";
    }

    [McpServerTool(Name = "todo_add_task_file_attachment_from_uploaded_content"), Description("Creates a file attachment from uploaded content without requiring manual base64 conversion. Supports plain text, base64, or data URL format.")]
    public async Task<string> AddTaskFileAttachmentFromUploadedContent(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Attachment file name, including extension.")] string name,
        [Description("Uploaded file content. Accepts plain text, base64, or data URL.")] string fileContent,
        [Description("Optional MIME type override. If omitted and input is a data URL, the data URL MIME type is used.")]
        string? contentType = null)
    {
        var bytes = DecodeFileContent(fileContent, out var detectedContentType);
        var finalContentType = string.IsNullOrWhiteSpace(contentType)
            ? (detectedContentType ?? "application/octet-stream")
            : contentType;

        var attachment = await graphService.AddTaskFileAttachmentAsync(listId, taskId, name, bytes, finalContentType);
        return $"Attachment created: '{attachment.Name}' (Id: {attachment.Id})";
    }

    [McpServerTool(Name = "todo_add_task_file_attachment_from_url"), Description("Downloads a file from an http/https URL and attaches it to a task. Uses delegated Graph download for SharePoint/OneDrive links and HTTP fallback for public URLs.")]
    public async Task<string> AddTaskFileAttachmentFromUrl(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Absolute http/https URL to download.")] string fileUrl,
        [Description("Optional attachment file name. If omitted, derived from URL path.")] string? name = null,
        [Description("Optional MIME type override.")] string? contentType = null)
    {
        var attachment = await graphService.AddTaskFileAttachmentFromUrlAsync(listId, taskId, fileUrl, name, contentType);
        return $"Attachment created from URL: '{attachment.Name}' (Id: {attachment.Id})";
    }

    [McpServerTool(Name = "todo_add_task_file_attachment_auto"), Description("Adds a task attachment from the best available input in this order: uploaded file content, file URL, then base64 content.")]
    public async Task<string> AddTaskFileAttachmentAuto(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Attachment file name, including extension.")] string name,
        [Description("Preferred when available: uploaded file content (plain text, base64, or data URL).")]
        string? fileContent = null,
        [Description("Preferred when upload bytes are unavailable: absolute http/https URL.")]
        string? fileUrl = null,
        [Description("Fallback: explicit base64 content.")]
        string? base64Content = null,
        [Description("Optional MIME type override.")]
        string? contentType = null)
    {
        if (!string.IsNullOrWhiteSpace(fileContent))
        {
            var bytes = DecodeFileContent(fileContent, out var detectedContentType);
            var finalContentType = string.IsNullOrWhiteSpace(contentType)
                ? (detectedContentType ?? "application/octet-stream")
                : contentType;

            var attachment = await graphService.AddTaskFileAttachmentAsync(listId, taskId, name, bytes, finalContentType);
            return $"Attachment created from uploaded content: '{attachment.Name}' (Id: {attachment.Id})";
        }

        if (!string.IsNullOrWhiteSpace(fileUrl))
        {
            var attachment = await graphService.AddTaskFileAttachmentFromUrlAsync(listId, taskId, fileUrl, name, contentType);
            return $"Attachment created from URL: '{attachment.Name}' (Id: {attachment.Id})";
        }

        if (!string.IsNullOrWhiteSpace(base64Content))
        {
            var bytes = Convert.FromBase64String(base64Content);
            var attachment = await graphService.AddTaskFileAttachmentAsync(
                listId,
                taskId,
                name,
                bytes,
                string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType);
            return $"Attachment created from base64 content: '{attachment.Name}' (Id: {attachment.Id})";
        }

        return "No attachment input provided. Pass one of: fileContent, fileUrl, or base64Content.";
    }

    [McpServerTool(Name = "todo_delete_task_attachment"), Description("Deletes a task attachment by attachmentId.")]
    public async Task<string> DeleteTaskAttachment(
        [Description("Microsoft Graph To Do list ID.")] string listId,
        [Description("Microsoft Graph task ID.")] string taskId,
        [Description("Attachment ID.")] string attachmentId)
    {
        await graphService.DeleteTaskAttachmentAsync(listId, taskId, attachmentId);
        return $"Attachment '{attachmentId}' deleted from task '{taskId}'.";
    }

    // ---------------------------------------------------------------------
    // Identity/debug tool
    // ---------------------------------------------------------------------

    [McpServerTool(Name = "todo_who_am_i"), Description("Returns the identity of the currently authenticated user from the Bearer token.")]
    public string WhoAmI(IHttpContextAccessor httpContextAccessor)
    {
        var claims = httpContextAccessor.HttpContext?.User?.Claims?.ToList();
        var name = claims?.FirstOrDefault(c => c.Type == "name")?.Value
                 ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Name)?.Value
                 ?? "unknown";
        var email = claims?.FirstOrDefault(c => c.Type == "preferred_username")?.Value
                 ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Upn)?.Value
                 ?? claims?.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value
                 ?? "unknown";
        var oid = claims?.FirstOrDefault(c => c.Type == "oid")?.Value
                 ?? claims?.FirstOrDefault(c => c.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value
                 ?? "unknown";

        return $"Authenticated as: {name} | Email: {email} | OID: {oid}";
    }

    private static byte[] DecodeFileContent(string fileContent, out string? detectedContentType)
    {
        detectedContentType = null;

        if (string.IsNullOrWhiteSpace(fileContent))
            return [];

        // Supports data URL format: data:<mime>;base64,<payload>
        if (fileContent.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            var commaIndex = fileContent.IndexOf(',');
            if (commaIndex > 5)
            {
                var header = fileContent[..commaIndex];
                var payload = fileContent[(commaIndex + 1)..];

                var semicolonIndex = header.IndexOf(';');
                if (semicolonIndex > 5)
                    detectedContentType = header[5..semicolonIndex];
                else if (header.Length > 5)
                    detectedContentType = header[5..];

                if (header.Contains(";base64", StringComparison.OrdinalIgnoreCase))
                    return Convert.FromBase64String(payload);

                return Encoding.UTF8.GetBytes(Uri.UnescapeDataString(payload));
            }
        }

        // Try base64 first. If not valid base64, treat as plain text.
        try
        {
            return Convert.FromBase64String(fileContent);
        }
        catch (FormatException)
        {
            return Encoding.UTF8.GetBytes(fileContent);
        }
    }
}