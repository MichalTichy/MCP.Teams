using TeamsMcp.Authentication;

namespace TeamsMcp.Tests.Support;

public sealed class TestGraphTokens : IGraphTokenProvider
{
    public Task<string> GetDelegatedTokenAsync(CancellationToken cancellationToken = default) => Task.FromResult("test-delegated");
    public Task<string> GetApplicationTokenAsync(CancellationToken cancellationToken = default) => throw new InvalidOperationException("Application credentials must not be used for file sending.");
}
