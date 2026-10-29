namespace TeamsMcp.Authentication;

public interface IGraphTokenProvider
{
    Task<string> GetDelegatedTokenAsync(CancellationToken cancellationToken = default);
    Task<string> GetApplicationTokenAsync(CancellationToken cancellationToken = default);
}
