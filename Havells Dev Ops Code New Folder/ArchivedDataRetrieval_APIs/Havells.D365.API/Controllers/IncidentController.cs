using Havells.D365.API.Contracts;
using Microsoft.AspNetCore.Mvc;
using Havells.D365.Entities.Incident.Response;
using Havells.D365.Services.Abstract;


namespace Havells.D365.API.Controllers
{
    [ApiController]
    public class IncidentController : ControllerBase
    {
        private readonly IIncidentRepository _incident;
        public IncidentController(IIncidentRepository _incident)
        {
            this._incident = _incident;
        }

        [HttpGet(ApiRoutes.Incident.GetIndidentByWorkOrderId)]
        public IActionResult GetIncidentByWorkOrderID(string workorderId)
        {
            if (string.IsNullOrEmpty(workorderId))
                return BadRequest(new IncidentResponse { Error = "Workorder ID Is Mandatory" });
            
            return Ok(_incident.GetIncidentByWorkOrderID(workorderId));
        }

        [HttpGet(ApiRoutes.Incident.GetIncidentById)]
        public IActionResult GetIncidentById(string incidentId)
        {
            if (string.IsNullOrEmpty(incidentId))
                return BadRequest(new IncidentResponse { Error = "Workorder ID Is Mandatory" });

            return Ok(_incident.GetIncidentById(incidentId));
        }
    }
}
