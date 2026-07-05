using Microsoft.AspNetCore.Mvc;
using O365C.Copilot.IncidentReporterSkill.Models;
using O365C.Copilot.IncidentReporterSkill.Services;

namespace O365C.Copilot.IncidentReporterSkill.Controllers;

[ApiController]
[Route("api/tickets")]
public class TicketsController : ControllerBase
{
    private readonly ITicketService _tickets;

    public TicketsController(ITicketService tickets) => _tickets = tickets;

    /// <summary>GET /api/tickets — open in browser on stage to show tickets accumulating live.</summary>
    [HttpGet]
    public async Task<IActionResult> GetAll(CancellationToken ct)
    {
        var tickets = await _tickets.GetAllTicketsAsync(ct);
        return Ok(tickets);
    }

    /// <summary>POST /api/tickets — seed test data without going through the bot conversation.</summary>
    [HttpPost]
    public async Task<IActionResult> Create([FromBody] CreateTicketRequest request, CancellationToken ct)
    {
        var ticket = await _tickets.CreateTicketAsync(request, ct);
        return CreatedAtAction(nameof(GetById), new { id = ticket.Id }, ticket);
    }

    /// <summary>GET /api/tickets/{id}</summary>
    [HttpGet("{id}")]
    public async Task<IActionResult> GetById(string id, CancellationToken ct)
    {
        var ticket = await _tickets.GetTicketByIdAsync(id, ct);
        return ticket is null ? NotFound() : Ok(ticket);
    }

    /// <summary>PATCH /api/tickets/{id}/status — resolve a ticket live from the audience.</summary>
    [HttpPatch("{id}/status")]
    public async Task<IActionResult> UpdateStatus(string id, [FromBody] UpdateStatusRequest body, CancellationToken ct)
    {
        try
        {
            var ticket = await _tickets.UpdateTicketStatusAsync(id, body.Status, ct);
            return Ok(ticket);
        }
        catch (KeyNotFoundException)
        {
            return NotFound();
        }
    }
}

public record UpdateStatusRequest(string Status);
