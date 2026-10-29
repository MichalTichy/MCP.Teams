using MediatR;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.ChatFiles.Notifications;
using TeamsMcp.ChatFiles.Usecases;
using TeamsMcp.ChatFiles.Storage;
using TeamsMcp.Configuration;
using TeamsMcp.Services;
using TeamsMcp.Tools;

namespace TeamsMcp.Tests.Support;

public sealed class ChatFilesTestContext : IDisposable
{
    public RecordingHttpHandler GraphHttp { get; } = new();
    public RecordingHttpHandler FileHttp { get; } = new();
    public List<ChatFilesMessageSentNotification> Notifications { get; } = [];
    private readonly ServiceProvider _services;
    private readonly HttpClient _http;
    private readonly HttpClient _fileHttp;
    public ChatFileTools Tool => new(_services.GetRequiredService<ISender>());
    public ChatTools Chats => new(_services.GetRequiredService<GraphClient>());

    public ChatFilesTestContext()
    {
        _http = new HttpClient(GraphHttp);
        _fileHttp = new HttpClient(FileHttp);
        var services = new ServiceCollection();
        var options = Options.Create(new TeamsOptions { GraphBaseUrl = "https://graph.example.com/v1.0", TenantId = "tenant-1" });
        services.AddSingleton(options);
        services.AddSingleton(new GraphClient(_http, new TestGraphTokens(),
            options));
        services.AddSingleton(new ChatFileSourceReader(_fileHttp, options));
        services.AddTransient<OneDriveChatFileStorage>();
        services.AddMediatR(config => config.RegisterServicesFromAssemblyContaining<SendChatFilesUsecase>());
        services.AddSingleton<INotificationHandler<ChatFilesMessageSentNotification>>(new RecordingSentNotificationHandler(Notifications));
        _services = services.BuildServiceProvider();
    }

    public Task<System.Text.Json.Nodes.JsonNode?> SendAsync(string message = "Report", params ChatFileReference[] files) =>
        _services.GetRequiredService<ISender>().Send(new SendChatFilesUsecase("chat-1", message,
            files.Length == 0 ? [Reference()] : files));

    public static ChatFileReference Reference() => new()
    {
        Id = "153fa47d-18c9-4179-be08-9879815a9f90",
        ContentUrl = "https://tenant.sharepoint.com/personal/me/Documents/Notino/report.xlsx",
        Name = "Výkaz říjen.xlsx"
    };

    public void Dispose()
    {
        _services.Dispose();
        _http.Dispose();
        _fileHttp.Dispose();
    }
}
