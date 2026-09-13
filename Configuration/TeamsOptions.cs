namespace TeamsMcp.Configuration;

public sealed class TeamsOptions
{
    public const string SectionName = "TeamsMcp";

    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string AdminPassword { get; set; } = string.Empty;
    public string DataPath { get; set; } = "/app/data";
    public string GraphBaseUrl { get; set; } = "https://graph.microsoft.com/v1.0";

    public string[] DelegatedScopes { get; set; } =
    [
        "User.Read",
        "Chat.ReadWrite",
        "ChatMessage.Send",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "ChannelMessage.Read.All",
        "ChannelMessage.Send",
        "Calendars.ReadWrite"
    ];

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(TenantId)) throw new InvalidOperationException("TeamsMcp:TenantId is required.");
        if (string.IsNullOrWhiteSpace(ClientId)) throw new InvalidOperationException("TeamsMcp:ClientId is required.");
        if (string.IsNullOrWhiteSpace(AdminPassword)) throw new InvalidOperationException("TeamsMcp:AdminPassword is required.");
    }
}
