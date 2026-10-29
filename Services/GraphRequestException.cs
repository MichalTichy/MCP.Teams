using System.Net;

namespace TeamsMcp.Services;

public sealed class GraphRequestException(HttpStatusCode statusCode, string message)
    : InvalidOperationException(message)
{
    public HttpStatusCode StatusCode { get; } = statusCode;
}
