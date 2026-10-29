using System.ComponentModel;
using System.Text.Json.Serialization;

namespace TeamsMcp.ChatFiles.Models;

public sealed record ChatFileReference
{
    [JsonPropertyName("id")]
    [Description("SharePoint file GUID returned by the separate storage tool, from sharepointIds.listItemUniqueId or the driveItem eTag. This is not the opaque driveItem ID.")]
    public required string Id { get; init; }

    [JsonPropertyName("contentUrl")]
    [Description("HTTPS webDavUrl of the existing OneDrive/SharePoint file, returned by the storage tool. All intended recipients must already have access.")]
    public required string ContentUrl { get; init; }

    [JsonPropertyName("name")]
    [Description("Display filename, including its extension, returned by the storage tool.")]
    public required string Name { get; init; }
}
