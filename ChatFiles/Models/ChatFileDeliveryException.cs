namespace TeamsMcp.ChatFiles.Models;

public sealed class ChatFileDeliveryException(
    IReadOnlyList<ChatFileReference> attachments,
    string messageOutcome,
    Exception innerException)
    : InvalidOperationException("The chat send did not complete normally. Inspect the chat before retrying; do not retry automatically.", innerException)
{
    public IReadOnlyList<ChatFileReference> Attachments { get; } = attachments;
    public string MessageOutcome { get; } = messageOutcome;
}
