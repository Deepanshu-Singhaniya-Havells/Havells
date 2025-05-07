using Havells.D365.API.Contracts;
using Havells.D365.Entities.ActivityDetails.Response;
using Havells.D365.Entities.Incident.Response;
using Havells.D365.Entities.POStatusTracker.Response;
using Havells.D365.Entities.ProductReqHeader.Response;
using Havells.D365.Entities.ProductRequest.Response;
using Havells.D365.Entities.WorkorderProduct.Response;
using Havells.D365.Entities.WorkOrders.Response;
using Havells.D365.Entities.WorkOrderService.Response;
using Havells.D365.Services.Abstract;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Runtime.Serialization.Json;

namespace Havells.D365.API.Controllers
{

    [ApiController]
    public class WorkOrdersController : ControllerBase
    {
        private readonly IWorkOrderRepository _workOrder;
        public WorkOrdersController(IWorkOrderRepository _workOrder)
        {
            this._workOrder = _workOrder;
        }

        [HttpGet(ApiRoutes.WorkOrders.GetWorkOrderByCustRef)]
        public IActionResult GetWorkOrderByCustRef(string custRef, int pageNumber, int recordsPerPage)
        {
            if (string.IsNullOrEmpty(custRef))
                return BadRequest(new WorkOrderResponse { Error = "Customer Reference Is Mandatory" });
            if(pageNumber==0)
                return BadRequest(new WorkOrderResponse { Error = "Page Number Is Mandatory" });

            if (pageNumber == 0)
                return BadRequest(new WorkOrderResponse { Error = "Records Per Page Is Mandatory" });

            return Ok(_workOrder.GetWorkOrderByCustRef(custRef,pageNumber,recordsPerPage));
        }

        [HttpGet(ApiRoutes.WorkOrders.GetWorkOrdersByName)]
        public IActionResult GetWorkOrdersByName(string workOrderName)
        {
            if (string.IsNullOrEmpty(workOrderName))
                return BadRequest(new WorkOrderResponse { Error = "Work Order Name Is Mandatory" });
            return Ok(_workOrder.GetWorkOrderByName(workOrderName));
        }

        [HttpGet(ApiRoutes.WorkOrders.GetWorkOrdersById)]
        public IActionResult GetWorkOrdersById(string workOrderId)
        {
            if (string.IsNullOrEmpty(workOrderId))
                return BadRequest(new WorkOrderResponse { Error = "workOrderId Is Mandatory" });
            return Ok(_workOrder.GetWorkOrderById(workOrderId));
        }

        [HttpGet(ApiRoutes.CommonEntityApi.GetD365ArchivedData)]
        public IActionResult GetD365ArchivedData(string RefId,int entityNo,int ApiType)
        {
            //entityNo=1 then Incident 
            //entityNo=2 then WorkOrderProduct
            //entityNo=3 then WorkOrderService
            //entityNo=4 then WorkOrder
            //entityNo=5 then POStatusTracker
            //entityNo=6 then ProductRequest
            //entityNo=7 then ProductRequestHeader
            //entityNo=8 then ActivityDetails for(Email,SMS,Phone,Task)
            //entityNo=9 then SAWActivity
            //entityNo=10 then SAWActivityApproval  
            //entityNo=11 then ClimLine
            //entityNo=12 then ClaimHeader

            //entityNo == 4 && ApiType == 4 then Jobsheet
            //entityNo == 4 && ApiType == 5 then JobPMSCount
            //entityNo == 5 && ApiType == 1 then GetPOStatusTrackerByID
            //entityNo == 5 && ApiType == 2 then GetPOStatusTrackerByJobID
            //entityNo == 6 && ApiType == 1 then GetProductRequestByID
            //entityNo == 6 && ApiType == 2 then GetProductRequestByJobID
            //entityNo == 7 && ApiType == 1 then GetProductReqHeaderByID
            //entityNo == 7 && ApiType == 2 then GetProductReqHeaderByJobID
            //entityNo == 8 && ApiType == 1 then GetActivityDetailsById
            //entityNo == 8 && ApiType == 2 then GetActivityDetailsByJobId
            //entityNo == 9 && ApiType == 1 then GetActivityById
            //entityNo == 9 && ApiType == 2 then GetActivityByJobId
            //entityNo == 10 && ApiType == 1 then GetActivityApprovalById
            //entityNo == 10 && ApiType == 2 then GetActivityApprovalByJobId
            //entityNo == 10 && ApiType == 3 then GetActivityApprovalByActivityId
            //entityNo == 11 && ApiType == 1 then GetClaimlineById
            //entityNo == 11 && ApiType == 2 then GetClaimlineByJobId
            //entityNo == 12 && ApiType == 1 then GetClaimheaderById

            if (entityNo == 1 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  // RefId means workOrderId
                    return BadRequest(new IncidentResponse { Error = "Workorder ID Is Mandatory" });

                return Ok(_workOrder.GetIncidentByWorkOrderID(RefId));
            }
            else if (entityNo == 1 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))       //here RefId means incidentId
                    return BadRequest(new IncidentResponse { Error = "Workorder ID Is Mandatory" });
                return Ok(_workOrder.GetIncidentById(RefId));
            }
            else if (entityNo == 2 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))    //here RefId means WorkOrderProductID
                    return BadRequest(new WorkOrderProductResponse { Error = "Workorder Product ID Is Mandatory" });

                return Ok(_workOrder.GetWOProductDetailByID(RefId));
            }
            else if (entityNo == 2 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))   //here workOrderId means WorkOrderIncidentID 
                    return BadRequest(new WorkOrderProductResponse { Error = "Incident ID Is Mandatory" });

                return Ok(_workOrder.GetWOProductDetailByIncidentID(RefId));
            }
            else if (entityNo == 2 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))  //here RefId means WorkOrderID
                    return BadRequest(new WorkOrderProductResponse { Error = "WorkOrder ID Is Mandatory" });

                return Ok(_workOrder.GetWOProductDetailByJobID(RefId));
            }
            else if(entityNo==3 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //here RefId means WorkOrderID
                    return BadRequest(new WorkOrderServiceResponse { Error = "Workorder ID Is Mandatory" });

                return Ok(_workOrder.GetWOServiceDetailByJobID(RefId));
            }
            else if(entityNo == 3 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //here RefId means WorkOrderServiceID  
                    return BadRequest(new WorkOrderServiceResponse { Error = "WorkOrder Service ID Is Mandatory" });
                return Ok(_workOrder.GetWOServiceDetailByID(RefId));
            }
            else if (entityNo == 3 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))   // here RefId means WorkOrderIncidentID
                    return BadRequest(new WorkOrderServiceResponse { Error = "WorkOrder Incident ID Is Mandatory" });
                return Ok(_workOrder.GetWOServiceDetailByIncidentID(RefId));
            }
            else if (entityNo == 4 && ApiType==1)
            {
                if (string.IsNullOrEmpty(RefId))  // Refid means workOrderName
                    return BadRequest(new WorkOrderResponse { Error = "Work Order Name Is Mandatory" });
                return Ok(_workOrder.GetWorkOrderByName(RefId));
            }
            else if (entityNo == 4 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means workOrderId 
                    return BadRequest(new WorkOrderResponse { Error = "workOrderId Is Mandatory" });
                return Ok(_workOrder.GetWorkOrderById(RefId));
            }
            else if (entityNo==4 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means custRef 
                    return BadRequest(new WorkOrderResponse { Error = "custRef Is Mandatory" });
                return Ok(_workOrder.GetWorkOrderByCustRef(RefId, 1, 100));
            }
            else if (entityNo == 4 && ApiType == 4)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobId 
                    return BadRequest(new JobSheetResponse { Error = "JobId Is Mandatory" });
                //var cc = _workOrder.GetJobSheetById(RefId);
                //DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(JobSheetResponse));
                //JobSheetResponse rootObject = (JobSheetResponse)ser.ReadObject(new MemoryStream(cc));
                //JsonConvert.DeserializeObject<JobSheetResponse>(cc);
                return Ok(_workOrder.GetJobSheetById(RefId));
            }
            else if (entityNo == 4 && ApiType == 5)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means CUSTOMERASSETID 
                    return BadRequest(new WorkOrderResponse { Error = "Customer Asset Id Is Mandatory" });
                return Ok(_workOrder.GetJobCountByCustomerAssetID(RefId));
            }
            else if (entityNo == 5 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means POStatusTrackerID 
                    return BadRequest(new POStatusTrackerResponse { Error = "POStatus Tracker Id Is Mandatory" });
                return Ok(_workOrder.GetPOStatusTrackerById(RefId));
            }
            else if (entityNo == 5 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID 
                    return BadRequest(new POStatusTrackerResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetPOStatusTrackerByJobId(RefId));
            }
            else if (entityNo == 5 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means PRHeaderNo
                    return BadRequest(new POStatusTrackerResponse { Error = "Product Request Header No Is Mandatory" });
                return Ok(_workOrder.GetPOStatusTrackerByPRHeader(RefId));
            }
            else if (entityNo == 6 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means ProductRequestID 
                    return BadRequest(new ProductRequestResponse { Error = "Product Request Id Is Mandatory" });
                return Ok(_workOrder.GetProductRequestById(RefId));
            }
            else if (entityNo == 6 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID
                    return BadRequest(new ProductRequestResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetProductRequestByJobId(RefId));
            }
            else if (entityNo == 6 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means PRHeaderNo  
                    return BadRequest(new ProductRequestResponse { Error = "Product Request Header No Is Mandatory" });
                return Ok(_workOrder.GetProductRequestByPRHeader(RefId));
            }

            else if (entityNo == 7 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means ProductRequestHeaderID 
                    return BadRequest(new ProductReqHeaderResponse { Error = "Product Request Header Id Is Mandatory" });
                return Ok(_workOrder.GetProductReqHeaderById(RefId));
            }
            else if (entityNo == 7 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID 
                    return BadRequest(new ProductReqHeaderResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetProductReqHeaderByJobId(RefId));
            }
            else if (entityNo == 8 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means ActivityID 
                    return BadRequest(new ActivityDetailsResponse { Error = "Activity Id Is Mandatory" });
                return Ok(_workOrder.GetActivityDetailsById(RefId));
            }
            else if (entityNo == 8 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID 
                    return BadRequest(new ActivityDetailsResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetActivityDetailsByJobId(RefId));
            }
            else if (entityNo == 9 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means SAWActivityID 
                    return BadRequest(new SAWActivityResponse { Error = "SAWActivity Id Is Mandatory" });
                return Ok(_workOrder.GetSAWActivityById(RefId));
            }
            else if (entityNo == 9 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID 
                    return BadRequest(new SAWActivityResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetSAWActivityByJobId(RefId));
            }
            else if (entityNo == 10 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means SAWActivityApprovalID 
                    return BadRequest(new SAWActivityApprovalResponse { Error = "SAWActivity Approval Id Is Mandatory" });
                return Ok(_workOrder.GetSAWActivityApprovalById(RefId));
            }
            else if (entityNo == 10 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID 
                    return BadRequest(new SAWActivityApprovalResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetSAWActivityApprovalByJobId(RefId));
            }
            else if (entityNo == 10 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means ActivityID 
                    return BadRequest(new SAWActivityApprovalResponse { Error = "SAWActivity Id Is Mandatory" });
                return Ok(_workOrder.GetSAWActivityApprovalByActivityId(RefId));
            }
            else if (entityNo == 11 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means SAWActivityApprovalID 
                    return BadRequest(new ClaimlineResponse { Error = "Claimline Id Is Mandatory" });
                return Ok(_workOrder.GetClaimlineById(RefId));
            }
            else if (entityNo == 11 && ApiType == 2)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means JobID 
                    return BadRequest(new ClaimlineResponse { Error = "Job Id Is Mandatory" });
                return Ok(_workOrder.GetClaimlineByJobId(RefId));
            }
            else if (entityNo == 11 && ApiType == 3)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means ClaimheaderID 
                    return BadRequest(new ClaimlineResponse { Error = "Claimheader Id Is Mandatory" });
                return Ok(_workOrder.GetClaimlineByClaimheaderId(RefId));
            }

            else if (entityNo == 12 && ApiType == 1)
            {
                if (string.IsNullOrEmpty(RefId))  //RefId  means SAWActivityApprovalID 
                    return BadRequest(new ClaimheaderResponse { Error = "Claimheader Id Is Mandatory" });
                return Ok(_workOrder.GetClaimheaderById(RefId));
            }


            else
            {
                return NotFound();
            }
        }
    }
}
