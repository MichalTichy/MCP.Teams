using System.Net;
using System.Text.Json.Nodes;
using ModelContextProtocol.Protocol;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.Tests.Support;
using TeamsMcp.Tools;

namespace TeamsMcp.Tests;

public sealed class ChatFileDeliveryTests
{
    [Fact]
    public async Task Existing_file_is_sent_as_a_reference_with_no_storage_or_sharing_operations()
    {
        using var context = new ChatFilesTestContext();
        var file = ChatFilesTestContext.Reference();
        var result = await context.SendAsync("<script>unsafe</script>\nPřehled", file);

        Assert.Equal("message-1", result?["id"]?.GetValue<string>());
        var message = Assert.Single(context.GraphHttp.Requests);
        Assert.Equal(HttpMethod.Post, message.Method);
        Assert.Equal("/v1.0/chats/chat-1/messages", message.Path);
        Assert.Equal("Bearer test-delegated", message.Authorization);
        Assert.Equal("reference", message.Json?["attachments"]?[0]?["contentType"]?.GetValue<string>());
        Assert.Equal(file.ContentUrl, message.Json?["attachments"]?[0]?["contentUrl"]?.GetValue<string>());
        Assert.Equal(file.Name, message.Json?["attachments"]?[0]?["name"]?.GetValue<string>());
        Assert.Contains("&lt;script&gt;unsafe&lt;/script&gt;<br>", message.Json?["body"]?["content"]?.GetValue<string>());
        Assert.Contains($"<attachment id=\"{file.Id}\">", message.Json?["body"]?["content"]?.GetValue<string>());
        Assert.Equal("message-1", Assert.Single(context.Notifications).MessageId);
    }

    [Fact]
    public async Task Multiple_existing_files_are_sent_together_without_being_downloaded()
    {
        using var context = new ChatFilesTestContext();
        var first = ChatFilesTestContext.Reference();
        var second = first with
        {
            Id = "668f7fa8-8129-4de7-b32b-fe1b442e6ef1", Name = "Faktura.pdf",
            ContentUrl = "https://tenant.sharepoint.com/personal/me/Documents/Notino/invoice.pdf"
        };
        await context.SendAsync("Reports", first, second);
        var message = Assert.Single(context.GraphHttp.Requests);
        Assert.Equal(2, message.Json?["attachments"]?.AsArray().Count);
        Assert.Equal(2, Assert.Single(context.Notifications).Files.Count);
    }

    [Fact]
    public async Task Send_timeout_returns_an_unknown_outcome_without_automatic_retry()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => throw new HttpRequestException("Simulated connection loss");
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", [ChatFilesTestContext.Reference()]);
        Assert.True(result.IsError);
        var details = JsonNode.Parse(Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text)!;
        Assert.Equal("unknown", details["messageOutcome"]?.GetValue<string>());
        Assert.Equal(ChatFilesTestContext.Reference().Id, details["attachments"]?[0]?["id"]?.GetValue<string>());
        Assert.Single(context.GraphHttp.Requests);
        Assert.Empty(context.Notifications);
    }

    [Fact]
    public async Task Missing_send_receipt_does_not_report_success()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { });
        var exception = await Assert.ThrowsAsync<ChatFileDeliveryException>(() => context.SendAsync());
        Assert.Equal("unknown", exception.MessageOutcome);
        Assert.Single(context.GraphHttp.Requests);
        Assert.Empty(context.Notifications);
    }

    [Fact]
    public async Task Microsoft_graph_rejection_is_returned_as_a_tool_error()
    {
        using var context = new ChatFilesTestContext();
        context.GraphHttp.Respond = _ => RecordingHttpHandler.Json(new { error = new { message = "Access denied" } }, HttpStatusCode.Forbidden);
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", [ChatFilesTestContext.Reference()]);
        Assert.True(result.IsError);
        Assert.Contains("Access denied", Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text);
        Assert.Single(context.GraphHttp.Requests);
    }

    [Theory]
    [InlineData("http://tenant.sharepoint.com/file.xlsx")]
    [InlineData("https://user:password@tenant.sharepoint.com/file.xlsx")]
    [InlineData("https://tenant.sharepoint.com/file.xlsx#fragment")]
    [InlineData("/workspace/artifacts/report.xlsx")]
    [InlineData("file:///C:/report.xlsx")]
    [InlineData("data:application/octet-stream;base64,AA==")]
    public async Task Unsupported_reference_targets_are_rejected_before_sending(string url)
    {
        using var context = new ChatFilesTestContext();
        await Assert.ThrowsAsync<ArgumentException>(() => context.SendAsync("Report", ChatFilesTestContext.Reference() with { ContentUrl = url }));
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Theory]
    [InlineData("opaque-drive-item-id")]
    [InlineData("00000000-0000-0000-0000-000000000000")]
    [InlineData("\" onmouseover=\"alert(1)")]
    public async Task Invalid_attachment_ids_are_rejected_before_sending(string id)
    {
        using var context = new ChatFilesTestContext();
        await Assert.ThrowsAsync<ArgumentException>(() => context.SendAsync("Report", ChatFilesTestContext.Reference() with { Id = id }));
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Duplicate_file_guids_are_rejected_before_sending()
    {
        using var context = new ChatFilesTestContext();
        var file = ChatFilesTestContext.Reference();
        await Assert.ThrowsAsync<ArgumentException>(() => context.SendAsync("Report", file, file with { Id = file.Id.ToUpperInvariant() }));
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Invalid_later_reference_is_detected_before_any_send()
    {
        using var context = new ChatFilesTestContext();
        var file = ChatFilesTestContext.Reference();
        await Assert.ThrowsAsync<ArgumentException>(() => context.SendAsync("Report", file,
            file with { Id = "668f7fa8-8129-4de7-b32b-fe1b442e6ef1", Name = "bad\r\nname.xlsx" }));
        Assert.Empty(context.GraphHttp.Requests);
    }

    [Fact]
    public async Task Existing_plain_text_tool_keeps_its_original_wire_contract()
    {
        using var handler = new RecordingHttpHandler();
        using var http = new HttpClient(handler);
        var graph = new TeamsMcp.Services.GraphClient(http, new TestGraphTokens(),
            Microsoft.Extensions.Options.Options.Create(new TeamsMcp.Configuration.TeamsOptions { GraphBaseUrl = "https://graph.example.com/v1.0" }));
        await new ChatTools(graph).SendChatMessage("chat-1", "Existing message");
        var message = Assert.Single(handler.Requests);
        Assert.Equal("text", message.Json?["body"]?["contentType"]?.GetValue<string>());
        Assert.Null(message.Json?["attachments"]);
    }
}
