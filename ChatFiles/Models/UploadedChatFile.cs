namespace TeamsMcp.ChatFiles.Models;

public sealed record UploadedChatFile(string DriveItemId, string Name, string? WebUrl, bool Shared);
