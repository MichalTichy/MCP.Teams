using System.Reflection;
using ModelContextProtocol.Server;
using TeamsMcp.Tests.Support;
using TeamsMcp.Tools;

namespace TeamsMcp.Tests;

public sealed class ChatFileToolContractTests
{
    [Fact]
    public void Existing_references_and_new_file_inputs_are_exposed_in_the_actual_mcp_input_schema()
    {
        using var context = new ChatFilesTestContext();
        var method = typeof(ChatFileTools).GetMethod(nameof(ChatFileTools.SendChatMessageWithAttachments))!;
        var tool = McpServerTool.Create(method, context.Tool);
        Assert.Equal("teams_send_chat_message_with_attachments", tool.ProtocolTool.Name);
        var schema = tool.ProtocolTool.InputSchema;
        Assert.True(schema.GetProperty("properties").TryGetProperty("attachments", out _));
        Assert.Contains("\"contentUrl\"", schema.GetRawText());
        Assert.Contains("\"id\"", schema.GetRawText());
        Assert.Contains("\"name\"", schema.GetRawText());
        Assert.Contains("\"files\"", schema.GetRawText());
        Assert.Contains("\"agentFileUrl\"", schema.GetRawText());
        Assert.Contains("\"base64Content\"", schema.GetRawText());
        Assert.DoesNotContain("cancellationToken", schema.GetRawText());
        var plainText = typeof(ChatTools).GetMethod(nameof(ChatTools.SendChatMessage))!;
        Assert.Equal("teams_send_chat_message", plainText.GetCustomAttribute<McpServerToolAttribute>()?.Name);
    }

    [Fact]
    public async Task Invalid_tool_input_returns_an_mcp_error_without_network_side_effects()
    {
        using var context = new ChatFilesTestContext();
        var result = await context.Tool.SendChatMessageWithAttachments("chat-1", "Report", []);
        Assert.True(result.IsError);
        Assert.Empty(context.GraphHttp.Requests);
    }
}
