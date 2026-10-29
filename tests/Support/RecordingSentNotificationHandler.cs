using MediatR;
using TeamsMcp.ChatFiles.Notifications;

namespace TeamsMcp.Tests.Support;

public sealed class RecordingSentNotificationHandler(List<ChatFilesMessageSentNotification> notifications)
    : INotificationHandler<ChatFilesMessageSentNotification>
{
    public Task Handle(ChatFilesMessageSentNotification notification, CancellationToken cancellationToken)
    {
        notifications.Add(notification);
        return Task.CompletedTask;
    }
}
