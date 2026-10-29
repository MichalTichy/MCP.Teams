using System.ComponentModel;
using System.Text.Json.Serialization;

namespace TeamsMcp.ChatFiles.Models;

public sealed record ChatFileUpload
{
    [JsonPropertyName("name"), Description("Filename including its extension, without directory components.")]
    public required string Name { get; init; }

    [JsonPropertyName("contentType"), Description("File MIME type; defaults to application/octet-stream.")]
    public string ContentType { get; init; } = "application/octet-stream";

    [JsonPropertyName("agentFileUrl"), Description("Run-scoped Gateway artifact URL. Through Codex Gateway, ./artifacts/report.xlsx is also accepted and converted to a temporary URL before forwarding. Supply exactly one of agentFileUrl or base64Content.")]
    public string? AgentFileUrl { get; init; }

    [JsonPropertyName("base64Content"), Description("Base64 file bytes for small files (maximum 1 MiB decoded). Prefer agentFileUrl for generated or larger files.")]
    public string? Base64Content { get; init; }
}
