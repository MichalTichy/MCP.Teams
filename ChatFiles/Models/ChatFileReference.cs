using System.ComponentModel;
using System.Text.Json.Serialization;

namespace TeamsMcp.ChatFiles.Models;

public sealed record ChatFileReference
{
    [JsonPropertyName("id")]
    [Description("Existing SharePoint file GUID from sharepointIds.listItemUniqueId or the driveItem eTag. This is not the opaque driveItem ID. For new files use the files parameter instead.")]
    public required string Id { get; init; }

    [JsonPropertyName("contentUrl")]
    [Description("HTTPS webDavUrl of the existing OneDrive/SharePoint file, returned by the storage tool. All intended recipients must already have access.")]
    public required string ContentUrl { get; init; }

    [JsonPropertyName("name")]
    [Description("Display filename, including its extension, returned by the storage tool.")]
    public required string Name { get; init; }
}
