using System.Text.RegularExpressions;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TicketRouter.Api.Services;

// Add 'partial' to the main class declaration
public partial class GraphMailService(GraphServiceClient graph)
{
    private readonly GraphServiceClient _graph = graph;

    // Extracts TrackingID#<number> from subject line
    [GeneratedRegex(@"TrackingID#(\d{15,})")]
    private static partial Regex TrackingIdRegex();

    public static string? ExtractTrackingId(string? subject)
    {
        if (string.IsNullOrEmpty(subject))
            return null;
        var m = TrackingIdRegex().Match(subject);
        return m.Success ? m.Groups[1].Value : null;
    }

    // Create/Ensures folders Inbox/Cases/Active/<ticketId – shortName>
    public async Task<string> EnsureTicketFolderAsync(string ticketId, string? shortName = null)
    {
        // Get Inbox
        var inbox = await _graph.Me.MailFolders
            .GetAsync(q => q.QueryParameters.Filter = "displayName eq 'Inbox'");
        var inboxId = inbox?.Value?.FirstOrDefault()?.Id ?? throw new InvalidOperationException("Inbox not found.");


        // Create/Ensure Inbox/Cases
        var cases = await GetOrCreateChildAsync(inboxId, "Cases");

        // Get or create Active
        var active = await GetOrCreateChildAsync(cases.Id!, "Active");

        // Create/Ensure ticket folder
        var name = string.IsNullOrWhiteSpace(shortName) ? ticketId : $"{ticketId} – {shortName}";
        var ticket = await GetOrCreateChildAsync(active.Id!, name);
        
        return ticket.Id!;
    }

    // Upsert message rule for ticket
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
                SubjectContains = [ $"TrackingID#{ticketId}" ]
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

    // Move selected message to destination folder
    public async Task MoveSelectedAsync(string messageId, string destinationFolderId)
        => await _graph.Me.Messages[messageId].Move.PostAsync(new() { DestinationId = destinationFolderId });

    // Copy sent message to destination folder
    public async Task CopySentAsync(string messageId, string destinationFolderId)
        => await _graph.Me.Messages[messageId].Copy.PostAsync(new() { DestinationId = destinationFolderId });

    // Close ticket: move folder from Active to Closed and disable rule
    public async Task<string> CloseTicketAsync(string ticketId)
    {
        // Get Inbox
        var inbox = await _graph.Me.MailFolders
                .GetAsync(q => q.QueryParameters.Filter = "displayName eq 'Inbox'");
        var inboxId = inbox!.Value!.First().Id!;

        // Ensure Inbox/Cases
        var cases = await GetOrCreateChildAsync(inboxId, "Cases");

        // Ensure Cases/Active and Cases/Closed
        var children = await _graph.Me.MailFolders[cases.Id!].ChildFolders.GetAsync();
        var active = children!.Value!.FirstOrDefault(f => f.DisplayName == "Active")
            ?? await _graph.Me.MailFolders[cases.Id!].ChildFolders.PostAsync(new() { DisplayName = "Active" });
        var closed = children!.Value!.FirstOrDefault(f => f.DisplayName == "Closed")
            ?? await _graph.Me.MailFolders[cases.Id!].ChildFolders.PostAsync(new() { DisplayName = "Closed" });


        // Find ticket folder under Cases/Active
        var actKids = await _graph.Me.MailFolders[active!.Id!].ChildFolders.GetAsync();
        var ticketFolder = actKids!.Value!.FirstOrDefault(f => f.DisplayName!.StartsWith(ticketId))
            ?? throw new InvalidOperationException("Ticket folder not found under Cases/Active.");


        // Ensure same name under Cases/Closed
        var closedKids = await _graph.Me.MailFolders[closed!.Id!].ChildFolders.GetAsync();
        var newFolder = closedKids!.Value!.FirstOrDefault(f => f.DisplayName == ticketFolder.DisplayName)
            ?? await _graph.Me.MailFolders[closed.Id!].ChildFolders.PostAsync(new() { DisplayName = ticketFolder.DisplayName });


        // Move messages (Top=100; consider pagination if needed)
        var msgs = await _graph.Me.MailFolders[ticketFolder.Id!].Messages.GetAsync(q => q.QueryParameters.Top = 100);
        if (msgs?.Value != null)
        {
            foreach (Message m in msgs.Value)
            {
                if (m?.Id != null && newFolder?.Id != null)
                    await MoveSelectedAsync(m.Id!, newFolder.Id!);
            }
        }

        // Disable rule
        var rules = await _graph.Me.MailFolders["inbox"].MessageRules.GetAsync();
        var rule = rules!.Value!.FirstOrDefault(r => r.DisplayName == $"TKT-{ticketId}");
        if (rule != null)
        {
            await _graph.Me.MailFolders["inbox"].MessageRules[rule.Id!]
                .PatchAsync(new MessageRule { IsEnabled = false });
        }

        return "Ticket closed and folder moved to Cases/Closed.";
    }

    // Helper to get or create child folder by name
    private async Task<MailFolder> GetOrCreateChildAsync(string parentId, string name)
    {
        MailFolderCollectionResponse? kids = await _graph.Me.MailFolders[parentId].ChildFolders
            .GetAsync(q => q.QueryParameters.Filter = $"displayName eq '{name.Replace("'", "''")}'");
        if (kids?.Value?.Count > 0) return kids.Value.First();
        var created = await _graph.Me.MailFolders[parentId].ChildFolders.PostAsync(new() { DisplayName = name });
        if (created is null)
            throw new InvalidOperationException($"Failed to create mail folder '{name}' under parent '{parentId}'.");
        return created!;
    }

    // Patch rule state (enable/disable)
    public async Task PatchRuleStateAsync(string ticketId, bool enable)
    {
        var rules = await _graph.Me.MailFolders["inbox"].MessageRules.GetAsync();
        var existing = rules?.Value?.FirstOrDefault(r => r.DisplayName == $"TKT-{ticketId}");
        if (existing is null) return;

        await _graph.Me.MailFolders["inbox"].MessageRules[existing.Id!]
            .PatchAsync(new MessageRule { IsEnabled = enable });
    }

}