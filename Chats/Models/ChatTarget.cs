using System.Text.Json.Nodes;
using TeamsMcp.Services;

namespace TeamsMcp.Chats.Models;

public static class ChatTarget
{
    public const string SelfId = "48:notes";
    public const string ParameterDescription =
        "Teams chat ID, or 'self' to message the connected Microsoft user privately (48:notes). " +
        "Sending to yourself needs no chat link or participant lookup.";

    public static string Resolve(string chatId)
    {
        var value = GraphClient.Require(chatId, nameof(chatId));
        return string.Equals(value, "self", StringComparison.OrdinalIgnoreCase) ? SelfId : value;
    }

    public static bool IsSelf(string chatId) => Resolve(chatId) == SelfId;

    public static JsonObject DescribeSelf() => new()
    {
        ["id"] = SelfId,
        ["isSelfChat"] = true,
        ["description"] = "Private messages to the connected Microsoft user. Use chatId 'self' or '48:notes'; no chat link is needed. This target is available independently of the recent-chat list."
    };
}
