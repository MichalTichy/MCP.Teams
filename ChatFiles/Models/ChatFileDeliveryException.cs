namespace TeamsMcp.ChatFiles.Models;

public sealed class ChatFileDeliveryException(
    IReadOnlyList<ChatFileReference> attachments,
    string messageOutcome,
    Exception innerException,
    IReadOnlyList<UploadedChatFile>? uploadedFiles = null)
    : InvalidOperationException("File delivery did not complete normally. Inspect messageOutcome and uploadedFiles before retrying; do not retry automatically.", innerException)
{
    public IReadOnlyList<ChatFileReference> Attachments { get; } = attachments;
    public string MessageOutcome { get; } = messageOutcome;
    public IReadOnlyList<UploadedChatFile> UploadedFiles { get; } = uploadedFiles ?? [];
}
