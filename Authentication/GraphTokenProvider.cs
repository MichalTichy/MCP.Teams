using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using TeamsMcp.Configuration;

namespace TeamsMcp.Authentication;

public sealed record DeviceCodePrompt(string UserCode, string VerificationUrl, DateTimeOffset ExpiresAt, string Message);
public sealed record ConnectionState(bool Connected, string? Username, DeviceCodePrompt? PendingPrompt, string? Error);

public sealed class GraphTokenProvider
{
    private readonly TeamsOptions _options;
    private readonly IPublicClientApplication _publicClient;
    private readonly IConfidentialClientApplication? _confidentialClient;
    private readonly string _cachePath;
    private readonly SemaphoreSlim _cacheLock = new(1, 1);
    private readonly object _stateLock = new();
    private Task? _deviceCodeTask;
    private Task<DeviceCodePrompt>? _promptTask;
    private DeviceCodePrompt? _pendingPrompt;
    private string? _lastError;

    public GraphTokenProvider(IOptions<TeamsOptions> options)
    {
        _options = options.Value;
        Directory.CreateDirectory(_options.DataPath);
        _cachePath = Path.Combine(_options.DataPath, "delegated-token-cache.bin");

        _publicClient = PublicClientApplicationBuilder.Create(_options.ClientId)
            .WithTenantId(_options.TenantId)
            .Build();
        RegisterPersistentCache(_publicClient.UserTokenCache);

        if (!string.IsNullOrWhiteSpace(_options.ClientSecret))
        {
            _confidentialClient = ConfidentialClientApplicationBuilder.Create(_options.ClientId)
                .WithTenantId(_options.TenantId)
                .WithClientSecret(_options.ClientSecret)
                .Build();
        }
    }

    public async Task<string> GetDelegatedTokenAsync(CancellationToken cancellationToken = default)
    {
        var account = (await _publicClient.GetAccountsAsync()).FirstOrDefault()
            ?? throw new InvalidOperationException("No Microsoft account is connected. Sign in through the Teams MCP admin page.");
        try
        {
            var result = await _publicClient.AcquireTokenSilent(_options.DelegatedScopes, account)
                .ExecuteAsync(cancellationToken);
            return result.AccessToken;
        }
        catch (MsalUiRequiredException exception)
        {
            throw new InvalidOperationException("The Microsoft sign-in must be renewed from the Teams MCP admin page.", exception);
        }
    }

    public async Task<string> GetApplicationTokenAsync(CancellationToken cancellationToken = default)
    {
        if (_confidentialClient is null)
            throw new InvalidOperationException("Teams call history requires TeamsMcp:ClientSecret.");
        var result = await _confidentialClient.AcquireTokenForClient(["https://graph.microsoft.com/.default"])
            .ExecuteAsync(cancellationToken);
        return result.AccessToken;
    }

    public async Task<DeviceCodePrompt> StartDeviceCodeSignInAsync(CancellationToken cancellationToken = default)
    {
        Task<DeviceCodePrompt> promptTask;
        lock (_stateLock)
        {
            if (_deviceCodeTask is { IsCompleted: false } && _promptTask is not null)
            {
                promptTask = _promptTask;
            }
            else
            {
                _pendingPrompt = null;
                _lastError = null;
                var promptReady = new TaskCompletionSource<DeviceCodePrompt>(TaskCreationOptions.RunContinuationsAsynchronously);
                _promptTask = promptReady.Task;
                _deviceCodeTask = CompleteDeviceCodeSignInAsync(promptReady, CancellationToken.None);
                promptTask = promptReady.Task;
            }
        }
        return await WaitForPromptAsync(promptTask, cancellationToken);
    }

    private async Task<DeviceCodePrompt> WaitForPromptAsync(Task<DeviceCodePrompt> promptTask, CancellationToken cancellationToken) =>
        await promptTask.WaitAsync(TimeSpan.FromSeconds(15), cancellationToken);

    private async Task CompleteDeviceCodeSignInAsync(TaskCompletionSource<DeviceCodePrompt> promptReady, CancellationToken cancellationToken)
    {
        try
        {
            await _publicClient.AcquireTokenWithDeviceCode(_options.DelegatedScopes, codeResult =>
            {
                var prompt = new DeviceCodePrompt(
                    codeResult.UserCode,
                    codeResult.VerificationUrl,
                    codeResult.ExpiresOn,
                    codeResult.Message);
                lock (_stateLock) _pendingPrompt = prompt;
                promptReady.TrySetResult(prompt);
                return Task.CompletedTask;
            }).ExecuteAsync(cancellationToken);
            lock (_stateLock) _pendingPrompt = null;
        }
        catch (Exception exception)
        {
            lock (_stateLock)
            {
                _lastError = exception.Message;
                _pendingPrompt = null;
            }
            promptReady.TrySetException(exception);
        }
    }

    public async Task<ConnectionState> GetConnectionStateAsync()
    {
        var account = (await _publicClient.GetAccountsAsync()).FirstOrDefault();
        lock (_stateLock)
            return new ConnectionState(account is not null, account?.Username, _pendingPrompt, _lastError);
    }

    public async Task DisconnectAsync()
    {
        foreach (var account in await _publicClient.GetAccountsAsync())
            await _publicClient.RemoveAsync(account);
        if (File.Exists(_cachePath)) File.Delete(_cachePath);
        lock (_stateLock)
        {
            _pendingPrompt = null;
            _lastError = null;
            _promptTask = null;
        }
    }

    private void RegisterPersistentCache(ITokenCache cache)
    {
        cache.SetBeforeAccessAsync(async args =>
        {
            await _cacheLock.WaitAsync();
            try
            {
                if (File.Exists(_cachePath))
                    args.TokenCache.DeserializeMsalV3(await File.ReadAllBytesAsync(_cachePath), shouldClearExistingCache: true);
            }
            finally { _cacheLock.Release(); }
        });
        cache.SetAfterAccessAsync(async args =>
        {
            if (!args.HasStateChanged) return;
            await _cacheLock.WaitAsync();
            try
            {
                var temporaryPath = _cachePath + ".tmp";
                await File.WriteAllBytesAsync(temporaryPath, args.TokenCache.SerializeMsalV3());
                File.Move(temporaryPath, _cachePath, overwrite: true);
            }
            finally { _cacheLock.Release(); }
        });
    }
}
