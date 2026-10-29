using System.Text;
using System.Text.Json.Nodes;

namespace TeamsMcp.Tests.Support;

public sealed record RecordedRequest(HttpMethod Method, Uri Url, byte[] Bytes, string? ContentType, string? Authorization)
{
    public string Path => Url.AbsolutePath;
    public JsonNode? Json => Bytes.Length == 0 ? null : JsonNode.Parse(Encoding.UTF8.GetString(Bytes));
}
