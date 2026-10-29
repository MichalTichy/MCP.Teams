using System.Net.Http.Headers;
using Microsoft.Extensions.Options;
using TeamsMcp.ChatFiles.Models;
using TeamsMcp.Configuration;

namespace TeamsMcp.ChatFiles.Storage;

public sealed class ChatFileSourceReader(HttpClient http, IOptions<TeamsOptions> options)
{
    public async Task<byte[]> ReadAsync(ChatFileUpload file, CancellationToken cancellationToken)
    {
        Validate(file);
        if (file.Base64Content is not null)
        {
            if (file.Base64Content.Length > 1_398_104)
                throw new ArgumentException("Inline files must be at most 1 MiB; use agentFileUrl for larger files.");
            byte[] bytes;
            try { bytes = Convert.FromBase64String(file.Base64Content); }
            catch (FormatException exception) { throw new ArgumentException("base64Content is invalid.", exception); }
            if (bytes.Length > 1024 * 1024) throw new ArgumentException("Inline files must be at most 1 MiB.");
            return bytes;
        }

        var url = new Uri(file.AgentFileUrl!, UriKind.Absolute);
        var maximum = options.Value.MaxUploadMegabytes * 1024L * 1024L;
        using var response = await http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        response.EnsureSuccessStatusCode();
        if (response.Content.Headers.ContentLength > maximum)
            throw new ArgumentException("The file exceeds the configured upload limit.");
        await using var source = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var output = new MemoryStream();
        var buffer = new byte[64 * 1024];
        int read;
        while ((read = await source.ReadAsync(buffer, cancellationToken)) != 0)
        {
            if (output.Length + read > maximum) throw new ArgumentException("The file exceeds the configured upload limit.");
            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
        }
        return output.ToArray();
    }

    public void Validate(ChatFileUpload file)
    {
        ArgumentNullException.ThrowIfNull(file);
        if (string.IsNullOrWhiteSpace(file.Name) || file.Name.Length > 200 ||
            file.Name.Any(character => char.IsControl(character) || "\\/:*?\"<>|".Contains(character)) ||
            file.Name.EndsWith('.') || file.Name.EndsWith(' '))
            throw new ArgumentException("A filename must contain 1-200 characters without directory separators or reserved characters.");
        if (!MediaTypeHeaderValue.TryParse(file.ContentType, out var mediaType) || mediaType.Parameters.Count != 0)
            throw new ArgumentException("contentType must be a valid MIME type without parameters.");
        if ((file.AgentFileUrl is null) == (file.Base64Content is null))
            throw new ArgumentException("Supply exactly one of agentFileUrl or base64Content.");
        if (file.AgentFileUrl is null) return;
        var allowed = new Uri(options.Value.GatewayFileBaseUrl);
        if (!Uri.TryCreate(file.AgentFileUrl, UriKind.Absolute, out var url) ||
            url.Scheme != allowed.Scheme || url.Authority != allowed.Authority ||
            url.UserInfo.Length != 0 || url.Query.Length != 0 || url.Fragment.Length != 0 ||
            !url.AbsolutePath.StartsWith("/_internal/mcp/artifacts/", StringComparison.Ordinal) ||
            url.AbsolutePath["/_internal/mcp/artifacts/".Length..] is not { Length: 43 } token ||
            token.Any(character => !char.IsAsciiLetterOrDigit(character) && character is not '-' and not '_'))
            throw new ArgumentException("agentFileUrl must be a run-scoped artifact URL from the configured Gateway origin.");
    }
}
