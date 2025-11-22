using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using TicketRouter.Api.Services;

namespace TicketRouter.Api.Controllers;

[ApiController]
[Route("api/[controller]")]
[Authorize] // requires the bearer token from Office.auth.getAccessToken
public class TicketsController(GraphMailService svc) : ControllerBase
{
    private readonly GraphMailService _svc = svc;

    public record RouteSelectedRequest(string MessageId, string Subject, string? ShortName);
    public record SentRouteRequest(string MessageId, string Subject);
    public record RuleStateRequest(string TicketId, bool Enable);
    public record CloseRequest(string TicketId);

    [HttpPost("routeSelected")]
    public async Task<IActionResult> RouteSelected([FromBody] RouteSelectedRequest dto)
    {
        var ticketId = GraphMailService.ExtractTrackingId(dto.Subject);
        if (ticketId is null) return BadRequest("No TrackingID found in subject.");
        var folderId = await _svc.EnsureTicketFolderAsync(ticketId, dto.ShortName);
        await _svc.UpsertRuleAsync(ticketId, folderId, enable: true);
        await _svc.MoveSelectedAsync(dto.MessageId, folderId);
        return Ok(new { ticketId, folderId, status = "routed" });
    }

    [HttpPost("sentRoute")]
    public async Task<IActionResult> SentRoute([FromBody] SentRouteRequest dto)
    {
        var ticketId = GraphMailService.ExtractTrackingId(dto.Subject);
        if (ticketId is null) return Ok(new { status = "no-ticket" });
        var folderId = await _svc.EnsureTicketFolderAsync(ticketId);
        await _svc.UpsertRuleAsync(ticketId, folderId, enable: true);
        await _svc.CopySentAsync(dto.MessageId, folderId);
        return Ok(new { ticketId, status = "copied" });
    }

    // ✅ Updated to only toggle rule state without changing its destination
    [HttpPost("state")]
    public async Task<IActionResult> SetRuleState([FromBody] RuleStateRequest dto)
    {
        await _svc.PatchRuleStateAsync(dto.TicketId, dto.Enable);   // <-- use the new method
        return Ok(new { status = dto.Enable ? "Rule enabled" : "Rule disabled" });
    }

    [HttpPost("close")]
    public async Task<IActionResult> Close([FromBody] CloseRequest dto)
    {
        var msg = await _svc.CloseTicketAsync(dto.TicketId);
        return Ok(new { status = msg });
    }
}
