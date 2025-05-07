using Havells.D365.API.Contracts;
using Microsoft.AspNetCore.Mvc;
using Havells.D365.Entities.WorkOrderService.Response;
using Havells.D365.Services.Abstract;

namespace Havells.D365.API.Controllers
{
   
    [ApiController]
    public class WorkOrderServiceController : ControllerBase
    {

        //public const string usp_GetWOServiceDetailByJobID = "[D365].[usp_GetWOServiceDetailByJobID]";
        //public const string usp_GetWOServiceDetailByID = "[D365].[usp_GetWOServiceDetailByID]";
        //public const string usp_GetWOServiceDetailByIncidentID = "[D365].[usp_GetWOServiceDetailByIncidentID]";
        private readonly IWOServiceRepository _service;
        public WorkOrderServiceController(IWOServiceRepository _service)
        {
            this._service = _service;
        }

        [HttpGet(ApiRoutes.WorkOrdService.GetWOServiceDetailByJobID)]
        public IActionResult GetWOServiceDetailByJobID(string WorkOrderID)
        {
            if (string.IsNullOrEmpty(WorkOrderID))
                return BadRequest(new WorkOrderServiceResponse { Error = "Workorder ID Is Mandatory" });

            return Ok(_service.GetWOServiceDetailByJobID(WorkOrderID));
        }

        [HttpGet(ApiRoutes.WorkOrdService.GetWOServiceDetailByID)]
        public IActionResult GetWOServiceDetailByID(string WorkOrderServiceID)
        {
            if (string.IsNullOrEmpty(WorkOrderServiceID))
                return BadRequest(new WorkOrderServiceResponse { Error = "WorkOrder Service ID Is Mandatory" });
            return Ok(_service.GetWOServiceDetailByID(WorkOrderServiceID));
        }

        [HttpGet(ApiRoutes.WorkOrdService.GetWOServiceDetailByIncidentID)]
        public IActionResult GetWOServiceDetailByIncidentID(string WorkOrderIncidentID)
        {
            if (string.IsNullOrEmpty(WorkOrderIncidentID))
                return BadRequest(new WorkOrderServiceResponse { Error = "WorkOrder Incident ID Is Mandatory" });
            return Ok(_service.GetWOServiceDetailByIncidentID(WorkOrderIncidentID));
        }

    }
}
