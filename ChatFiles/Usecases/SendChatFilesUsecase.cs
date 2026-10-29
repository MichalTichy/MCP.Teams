using System.Text.Json.Nodes;
using MediatR;
using TeamsMcp.ChatFiles.Models;

namespace TeamsMcp.ChatFiles.Usecases;

public sealed record SendChatFilesUsecase(string ChatId, string Message, IReadOnlyList<ChatFileReference> Attachments,
    IReadOnlyList<ChatFileUpload>? Files = null)
    : IRequest<JsonNode?>;
