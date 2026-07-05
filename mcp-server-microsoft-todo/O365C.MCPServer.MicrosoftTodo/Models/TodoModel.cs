namespace O365C.MCPServer.MicrosoftTodo.Models
{
    public class TodoList
    {
        public string Id { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public bool IsOwner { get; set; }
        public bool IsShared { get; set; }
    }

    public class TodoTask
    {
        public string Id { get; set; } = "";
        public string ListId { get; set; } = "";
        public string Title { get; set; } = "";
        public string Status { get; set; } = "";
        public string Importance { get; set; } = "";
        public DateTime? DueDateTime { get; set; }
        public DateTime? CompletedDateTime { get; set; }
        public DateTime CreatedDateTime { get; set; }
        public string? Body { get; set; }
    }

    public class TodoChecklistItem
    {
        public string Id { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public bool IsChecked { get; set; }
        public DateTime? CreatedDateTime { get; set; }
        public DateTime? CheckedDateTime { get; set; }
    }

    public class TodoLinkedResource
    {
        public string Id { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string? ApplicationName { get; set; }
        public string? ExternalId { get; set; }
        public string? WebUrl { get; set; }
    }

    public class TodoTaskAttachment
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string? ContentType { get; set; }
        public int? Size { get; set; }
        public DateTime? LastModifiedDateTime { get; set; }
        public string? AttachmentType { get; set; }
    }
}
