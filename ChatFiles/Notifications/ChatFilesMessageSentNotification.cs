using MediatR;
using TeamsMcp.ChatFiles.Models;

namespace TeamsMcp.ChatFiles.Notifications;

public sealed record ChatFilesMessageSentNotification(string ChatId, string MessageId, IReadOnlyList<ChatFileReference> Files)
    : INotification;
