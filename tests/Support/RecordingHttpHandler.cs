using System.Net;
using System.Text;
using System.Text.Json;

namespace TeamsMcp.Tests.Support;

public sealed class RecordingHttpHandler : HttpMessageHandler
{
    public List<RecordedRequest> Requests { get; } = [];
    public Func<RecordedRequest, HttpResponseMessage>? Respond { get; set; }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var recorded = new RecordedRequest(request.Method, request.RequestUri!,
            request.Content is null ? [] : await request.Content.ReadAsByteArrayAsync(cancellationToken),
            request.Content?.Headers.ContentType?.MediaType, request.Headers.Authorization?.ToString());
        Requests.Add(recorded);
        if (Respond is not null) return Respond(recorded);
        if (request.Method == HttpMethod.Post && recorded.Path.EndsWith("/messages"))
            return Json(new { id = "message-1" }, HttpStatusCode.Created);
        throw new InvalidOperationException("Unexpected API operation: " + request.Method + " " + request.RequestUri);
    }

    public static HttpResponseMessage Json(object value, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(value), Encoding.UTF8, "application/json") };
}
