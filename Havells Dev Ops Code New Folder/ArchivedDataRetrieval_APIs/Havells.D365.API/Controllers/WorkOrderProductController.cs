using Havells.D365.API.Contracts;
using Microsoft.AspNetCore.Mvc;
using Havells.D365.Entities.WorkorderProduct.Response;
using Havells.D365.Services.Abstract;


namespace Havells.D365.API.Controllers
{
    
    [ApiController]
    public class WorkOrderProductController : ControllerBase
    {
        private readonly IWorkOrderProductRepository _WOProduct;
        public WorkOrderProductController(IWorkOrderProductRepository _WOProduct)
        {
            this._WOProduct = _WOProduct;
        }

        [HttpGet(ApiRoutes.WorkOrderProduct.GetWOProductDetailByID)]
        public IActionResult GetWOProductDetailByID(string WorkOrderProductID)
        {
            if (string.IsNullOrEmpty(WorkOrderProductID))
                return BadRequest(new WorkOrderProductResponse { Error = "Workorder Product ID Is Mandatory" });

            return Ok(_WOProduct.GetWOProductDetailByID(WorkOrderProductID));
        }

        [HttpGet(ApiRoutes.WorkOrderProduct.GetWOProductDetailByIncidentID)]
        public IActionResult GetWOProductDetailByIncidentID(string WorkOrderIncidentID)
        {
            if (string.IsNullOrEmpty(WorkOrderIncidentID))
                return BadRequest(new WorkOrderProductResponse { Error = "Incident ID Is Mandatory" });

            return Ok(_WOProduct.GetWOProductDetailByIncidentID(WorkOrderIncidentID));
        }

        [HttpGet(ApiRoutes.WorkOrderProduct.GetWOProductDetailByJobID)]
        public IActionResult GetWOProductDetailByJobID(string WorkOrderID)
        {
            if (string.IsNullOrEmpty(WorkOrderID))
                return BadRequest(new WorkOrderProductResponse { Error = "WorkOrder ID Is Mandatory" });

            return Ok(_WOProduct.GetWOProductDetailByJobID(WorkOrderID));
        }
    }
}
