using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Options;
using TeamsMcp.Authentication;
using TeamsMcp.Configuration;

namespace TeamsMcp.Services;

public sealed class GraphClient(HttpClient httpClient, GraphTokenProvider tokens, IOptions<TeamsOptions> options)
{
    private readonly TeamsOptions _options = options.Value;

    public Task<JsonNode?> DelegatedAsync(HttpMethod method, string path, object? body = null, CancellationToken cancellationToken = default) =>
        SendAsync(method, path, body, tokens.GetDelegatedTokenAsync, cancellationToken);

    public Task<JsonNode?> ApplicationAsync(HttpMethod method, string path, object? body = null, CancellationToken cancellationToken = default) =>
        SendAsync(method, path, body, tokens.GetApplicationTokenAsync, cancellationToken);

    public Task<string> ApplicationTextAsync(
        HttpMethod method,
        string path,
        string accept,
        CancellationToken cancellationToken = default) =>
        SendTextAsync(method, path, accept, tokens.GetApplicationTokenAsync, cancellationToken);

    public Task<string> DelegatedTextAsync(
        HttpMethod method,
        string path,
        string accept,
        CancellationToken cancellationToken = default) =>
        SendTextAsync(method, path, accept, tokens.GetDelegatedTokenAsync, cancellationToken);

    private async Task<JsonNode?> SendAsync(
        HttpMethod method,
        string path,
        object? body,
        Func<CancellationToken, Task<string>> tokenFactory,
        CancellationToken cancellationToken)
    {
        var baseUrl = _options.GraphBaseUrl.TrimEnd('/');
        using var request = new HttpRequestMessage(method, path.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            ? path
            : baseUrl + "/" + path.TrimStart('/'));
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await tokenFactory(cancellationToken));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        if (body is not null)
            request.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        using var response = await httpClient.SendAsync(request, cancellationToken);
        var text = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var safeMessage = TryGetGraphError(text) ?? response.ReasonPhrase ?? "Microsoft Graph request failed";
            throw new InvalidOperationException($"Microsoft Graph returned {(int)response.StatusCode}: {safeMessage}");
        }
        return string.IsNullOrWhiteSpace(text) ? null : JsonNode.Parse(text);
    }

    private async Task<string> SendTextAsync(
        HttpMethod method,
        string path,
        string accept,
        Func<CancellationToken, Task<string>> tokenFactory,
        CancellationToken cancellationToken)
    {
        var baseUrl = _options.GraphBaseUrl.TrimEnd('/');
        using var request = new HttpRequestMessage(method, path.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            ? path
            : baseUrl + "/" + path.TrimStart('/'));
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await tokenFactory(cancellationToken));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(Require(accept, nameof(accept))));

        using var response = await httpClient.SendAsync(request, cancellationToken);
        var text = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var safeMessage = TryGetGraphError(text) ?? response.ReasonPhrase ?? "Microsoft Graph request failed";
            throw new InvalidOperationException($"Microsoft Graph returned {(int)response.StatusCode}: {safeMessage}");
        }

        return text;
    }

    public static string Print(JsonNode? node) => node?.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) ?? "{\n  \"success\": true\n}";
    public static string Id(string value) => Uri.EscapeDataString(Require(value, "id"));
    public static string Require(string value, string name) => !string.IsNullOrWhiteSpace(value) ? value.Trim() : throw new ArgumentException($"{name} is required.");
    public static int Top(int value, int maximum = 50) => Math.Clamp(value, 1, maximum);

    private static string? TryGetGraphError(string text)
    {
        try { return JsonNode.Parse(text)?["error"]?["message"]?.GetValue<string>(); }
        catch (JsonException) { return null; }
    }
}
