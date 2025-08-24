using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TicketRouter.Api.Services;

public class GraphMailService
{
    private readonly GraphServiceClient _graph;

    public GraphMailService(GraphServiceClient graph) => _graph = graph;

    public static string? ExtractTrackingId(string? subject)
    {
        if (string.IsNullOrEmpty(subject)) return null;
        var m = System.Text.RegularExpressions.Regex.Match(subject, @"TrackingID#(\d{15,})");
        return m.Success ? m.Groups[1].Value : null;
    }

    public async Task<string> EnsureTicketFolderAsync(string ticketId, string? shortName = null)
    {
        // Get Inbox
        var inbox = await _graph.Me.MailFolders
            .GetAsync(q => q.QueryParameters.Filter = "displayName eq 'Inbox'");
        var inboxId = inbox?.Value?.FirstOrDefault()?.Id ?? throw new InvalidOperationException("Inbox not found.");

        // Get or create Active
        var active = await GetOrCreateChildAsync(inboxId, "Active");

        // Create/Ensure ticket folder
        var name = string.IsNullOrWhiteSpace(shortName) ? ticketId : $"{ticketId} – {shortName}";
        var ticket = await GetOrCreateChildAsync(active.Id!, name);
        return ticket.Id!;
    }

    public async Task UpsertRuleAsync(string ticketId, string folderId, bool enable = true)
    {
        var rules = await _graph.Me.MailFolders["inbox"].MessageRules.GetAsync();
        var existing = rules?.Value?.FirstOrDefault(r => r.DisplayName == $"TKT-{ticketId}");

        var body = new MessageRule
        {
            DisplayName = $"TKT-{ticketId}",
            IsEnabled = enable,
            Conditions = new MessageRulePredicates
            {
                SubjectContains = new List<string> { $"TrackingID#{ticketId}" }
            },
            Actions = new MessageRuleActions
            {
                MoveToFolder = $"https://graph.microsoft.com/v1.0/me/mailFolders('{folderId}')",
                StopProcessingRules = true
            }
        };

        if (existing != null)
            await _graph.Me.MailFolders["inbox"].MessageRules[existing.Id!].PatchAsync(body);
        else
            await _graph.Me.MailFolders["inbox"].MessageRules.PostAsync(body);
    }

    public async Task MoveSelectedAsync(string messageId, string destinationFolderId)
        => await _graph.Me.Messages[messageId].Move.PostAsync(new() { DestinationId = destinationFolderId });

    public async Task CopySentAsync(string messageId, string destinationFolderId)
        => await _graph.Me.Messages[messageId].Copy.PostAsync(new() { DestinationId = destinationFolderId });

    public async Task<string> CloseTicketAsync(string ticketId)
    {
        // Find Inbox/Active/<ticket…> and recreate in Inbox/Closed
        var inbox = await _graph.Me.MailFolders.GetAsync(q => q.QueryParameters.Filter = "displayName eq 'Inbox'");
        var inboxId = inbox!.Value!.First().Id!;

        var children = await _graph.Me.MailFolders[inboxId].ChildFolders.GetAsync();
        var active = children!.Value!.FirstOrDefault(f => f.DisplayName == "Active")
                    ?? await _graph.Me.MailFolders[inboxId].ChildFolders.PostAsync(new() { DisplayName = "Active" });
        var closed = children!.Value!.FirstOrDefault(f => f.DisplayName == "Closed")
                    ?? await _graph.Me.MailFolders[inboxId].ChildFolders.PostAsync(new() { DisplayName = "Closed" });

        var actKids = await _graph.Me.MailFolders[active!.Id!].ChildFolders.GetAsync();
        var ticketFolder = actKids!.Value!.FirstOrDefault(f => f.DisplayName!.StartsWith(ticketId))
                           ?? throw new InvalidOperationException("Ticket folder not found under Active.");

        // Ensure same name under Closed
        var closedKids = await _graph.Me.MailFolders[closed!.Id!].ChildFolders.GetAsync();
        var newFolder = closedKids!.Value!.FirstOrDefault(f => f.DisplayName == ticketFolder.DisplayName)
                        ?? await _graph.Me.MailFolders[closed.Id!].ChildFolders.PostAsync(new() { DisplayName = ticketFolder.DisplayName });

        // Move all messages
        var msgs = await _graph.Me.MailFolders[ticketFolder.Id!].Messages.GetAsync(q => q.QueryParameters.Top = 100);
        foreach (var m in msgs!.Value!) await MoveSelectedAsync(m.Id!, newFolder.Id!);

        // Disable rule
        var rules = await _graph.Me.MailFolders["inbox"].MessageRules.GetAsync();
        var rule = rules!.Value!.FirstOrDefault(r => r.DisplayName == $"TKT-{ticketId}");
        if (rule != null)
            await _graph.Me.MailFolders["inbox"].MessageRules[rule.Id!]
                .PatchAsync(new MessageRule { IsEnabled = false });

        return "Ticket closed and folder moved to Closed.";
    }

    private async Task<MailFolder> GetOrCreateChildAsync(string parentId, string name)
    {
        var kids = await _graph.Me.MailFolders[parentId].ChildFolders
            .GetAsync(q => q.QueryParameters.Filter = $"displayName eq '{name.Replace("'", "''")}'");
        if (kids?.Value?.Any() == true) return kids.Value.First();
        return await _graph.Me.MailFolders[parentId].ChildFolders.PostAsync(new() { DisplayName = name })!;
    }
}