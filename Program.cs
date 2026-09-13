using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.Options;
using TeamsMcp.Authentication;
using TeamsMcp.Configuration;
using TeamsMcp.Services;
using TeamsMcp.Tools;

var builder = WebApplication.CreateBuilder(args);
var dataPath = builder.Configuration.GetSection(TeamsOptions.SectionName).GetValue<string>(nameof(TeamsOptions.DataPath))
    ?? "/app/data";
Directory.CreateDirectory(dataPath);
builder.Services.AddDataProtection()
    .PersistKeysToFileSystem(new DirectoryInfo(Path.Combine(dataPath, "data-protection")))
    .SetApplicationName("TeamsMcp");
builder.Services.AddOptions<TeamsOptions>()
    .Bind(builder.Configuration.GetSection(TeamsOptions.SectionName));
builder.Services.AddSingleton<GraphTokenProvider>();
builder.Services.AddHttpClient<GraphClient>();
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.Cookie.Name = "teams-mcp-admin";
        options.Cookie.HttpOnly = true;
        options.Cookie.SameSite = SameSiteMode.Lax;
        options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
        options.LoginPath = "/admin/login";
        options.ExpireTimeSpan = TimeSpan.FromHours(8);
    });
builder.Services.AddAuthorization();
builder.Services.AddMcpServer()
    .WithHttpTransport()
    .WithTools<ChatTools>()
    .WithTools<ChannelTools>()
    .WithTools<MeetingTools>()
    .WithTools<CallHistoryTools>();

var app = builder.Build();
app.Services.GetRequiredService<IOptions<TeamsOptions>>().Value.Validate();
app.UseAuthentication();
app.UseAuthorization();

app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));

app.MapGet("/admin/login", () => Results.Content(LoginPage(), "text/html"));
app.MapPost("/admin/login", async (HttpContext context, IOptions<TeamsOptions> options) =>
{
    var form = await context.Request.ReadFormAsync();
    if (!SecretsEqual(form["password"].ToString(), options.Value.AdminPassword))
        return Results.Content(LoginPage("The password is incorrect."), "text/html", statusCode: 401);

    var identity = new ClaimsIdentity(
        [new Claim(ClaimTypes.Name, "Samwise owner")], CookieAuthenticationDefaults.AuthenticationScheme);
    await context.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));
    return Results.Redirect("/admin");
}).DisableAntiforgery();

app.MapGet("/admin", async (GraphTokenProvider tokens) =>
    Results.Content(AdminPage(await tokens.GetConnectionStateAsync()), "text/html"))
    .RequireAuthorization();

app.MapPost("/admin/connect", async (GraphTokenProvider tokens, CancellationToken cancellationToken) =>
{
    await tokens.StartDeviceCodeSignInAsync(cancellationToken);
    return Results.Redirect("/admin");
}).DisableAntiforgery().RequireAuthorization();

app.MapPost("/admin/disconnect", async (GraphTokenProvider tokens) =>
{
    await tokens.DisconnectAsync();
    return Results.Redirect("/admin");
}).DisableAntiforgery().RequireAuthorization();

app.MapPost("/admin/logout", async (HttpContext context) =>
{
    await context.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.Redirect("/admin/login");
}).DisableAntiforgery().RequireAuthorization();

app.MapMcp();
app.Run();

static bool SecretsEqual(string supplied, string configured)
{
    var left = Encoding.UTF8.GetBytes(supplied);
    var right = Encoding.UTF8.GetBytes(configured);
    return left.Length == right.Length && CryptographicOperations.FixedTimeEquals(left, right);
}

const string Styles = """
:root{font-family:Inter,Segoe UI,sans-serif;color:#302820;background:#f5efe3}*{box-sizing:border-box}body{margin:0;min-height:100vh;background:radial-gradient(circle at 12% 15%,#efae3d26,transparent 28%),#f5efe3}main{min-height:100vh;display:grid;place-items:center;padding:24px}.card{width:min(430px,100%);padding:30px;border:1px solid #e5d8c7;border-radius:20px;background:#fffaf2;box-shadow:0 18px 55px #59411b1a}.wide{width:min(660px,100%)}.eyebrow{color:#a65f18;font-size:12px;font-weight:800;letter-spacing:.14em}h1{font-size:32px;margin:.25rem 0 .65rem}h2{font-size:18px}p{color:#716459;line-height:1.55}label{display:grid;gap:7px;font-size:13px;font-weight:650}input{padding:12px;border:1px solid #d9c9b6;border-radius:10px;background:#fff}button{border:0;border-radius:10px;padding:11px 16px;background:#df9223;color:#2f2417;font-weight:750;cursor:pointer}.secondary{background:#e9dfd0}.quiet{background:transparent;color:#6d5f51}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:22px}.status{display:grid;gap:3px;margin:20px 0;padding:14px;border-radius:12px;background:#eee5d9}.status.good{background:#e8efd8}.status span,.hint{font-size:13px}.device{margin:18px 0;padding:16px;border:1px solid #d9c8ae;border-radius:12px;background:#fcf5e8}.device code{display:inline-block;padding:8px 12px;border-radius:8px;background:#302820;color:#fff;font-size:22px;letter-spacing:.12em}.error{margin:12px 0;padding:10px;border-radius:9px;background:#f8dbd4;color:#8d2f20}a{color:#a65f18}
""";

static string LoginPage(string? error = null) => $$"""
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Teams MCP sign in</title><style>{{Styles}}</style></head><body><main><section class="card"><p class="eyebrow">SAMWISE · TEAMS MCP</p><h1>Administration</h1><p>Enter the private admin password to manage the connected Microsoft account.</p>
{{(error is null ? "" : $"<div class=\"error\">{HtmlEncoder.Default.Encode(error)}</div>")}}
<form method="post" action="/admin/login"><label>Password<input name="password" type="password" required autofocus></label><button>Continue</button></form></section></main></body></html>
""";

static string AdminPage(ConnectionState state)
{
    var prompt = state.PendingPrompt;
    var status = state.Connected
        ? $"<div class=\"status good\"><strong>Connected</strong><span>{HtmlEncoder.Default.Encode(state.Username ?? "Microsoft account")}</span></div>"
        : "<div class=\"status\"><strong>Not connected</strong><span>Chat, channel, and meeting tools cannot run yet.</span></div>";
    var promptMarkup = prompt is null ? "" : $$"""
<div class="device"><h2>Complete Microsoft sign-in</h2><p>Open <a href="{{HtmlEncoder.Default.Encode(prompt.VerificationUrl)}}" target="_blank" rel="noreferrer">{{HtmlEncoder.Default.Encode(prompt.VerificationUrl)}}</a> and enter:</p><code>{{HtmlEncoder.Default.Encode(prompt.UserCode)}}</code><p>This page can be refreshed after sign-in.</p></div>
""";
    var error = state.Error is null ? "" : $"<div class=\"error\">{HtmlEncoder.Default.Encode(state.Error)}</div>";
    return $$"""
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Teams MCP</title><style>{{Styles}}</style></head><body><main><section class="card wide"><p class="eyebrow">SAMWISE · TEAMS MCP</p><h1>Microsoft account</h1><p>One delegated account supplies Teams chats, channels, and calendar-backed Teams meetings. Call history separately uses this application's client credentials.</p>{{status}}{{error}}{{promptMarkup}}<div class="actions"><form method="post" action="/admin/connect"><button>{{(state.Connected ? "Renew sign-in" : "Connect Microsoft account")}}</button></form>{{(state.Connected ? "<form method=\"post\" action=\"/admin/disconnect\"><button class=\"secondary\">Disconnect</button></form>" : "")}}<form method="post" action="/admin/logout"><button class="quiet">Log out</button></form></div><p class="hint">The Azure app registration must allow public client flows. Grant delegated Teams/calendar permissions and the CallRecords.Read.All application permission described in the repository README.</p></section></main></body></html>
""";
}

public partial class Program;
