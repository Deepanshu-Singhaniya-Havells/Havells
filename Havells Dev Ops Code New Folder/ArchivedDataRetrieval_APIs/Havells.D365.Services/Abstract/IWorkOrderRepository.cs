using Havells.D365.Entities.WorkOrders.Response;
using Havells.D365.Entities.Incident.Response;
using Havells.D365.Entities.WorkorderProduct.Response;
using Havells.D365.Entities.WorkOrderService.Response;
using Havells.D365.Entities.POStatusTracker.Response;
using Havells.D365.Entities.ProductRequest.Response;
using Havells.D365.Entities.ProductReqHeader.Response;
using Havells.D365.Entities.ActivityDetails.Response;

namespace Havells.D365.Services.Abstract
{
    public interface IWorkOrderRepository
    {
        WorkOrderResponse GetWorkOrderByName(string workOrderName);
        WorkOrderResponse GetWorkOrderById(string workOrderId);
        WorkOrderResponse GetWorkOrderByCustRef(string custRef, int pageNumber, int recordsPerPage);
        IncidentResponse GetIncidentById(string incidentId);
        IncidentResponse GetIncidentByWorkOrderID(string workorderId);
        WorkOrderProductResponse GetWOProductDetailByID(string WorkOrderProductID);
        WorkOrderProductResponse GetWOProductDetailByIncidentID(string WorkOrderIncidentID);
        WorkOrderProductResponse GetWOProductDetailByJobID(string WorkOrderID);
        WorkOrderServiceResponse GetWOServiceDetailByJobID(string WorkOrderID);
        WorkOrderServiceResponse GetWOServiceDetailByID(string WorkOrderServiceID);
        WorkOrderServiceResponse GetWOServiceDetailByIncidentID(string WorkOrderIncidentID);
        JobSheetResponse GetJobSheetById(string JobId);
        JobSheetCountResponse GetJobCountByCustomerAssetID(string CustomerAssetID);
        POStatusTrackerResponse GetPOStatusTrackerById(string POStatusTrackerID);
        POStatusTrackerResponse GetPOStatusTrackerByJobId(string JobID);
        POStatusTrackerResponse GetPOStatusTrackerByPRHeader(string JobID);

        ProductRequestResponse GetProductRequestById(string ProductRequestID);
        ProductRequestResponse GetProductRequestByJobId(string JobID);
        ProductRequestResponse GetProductRequestByPRHeader(string JobID);

        ProductReqHeaderResponse GetProductReqHeaderById(string ProductReqHeaderID);
        ProductReqHeaderResponse GetProductReqHeaderByJobId(string JobID);

        ActivityDetailsResponse GetActivityDetailsById(string ActivityID);
        ActivityDetailsResponse GetActivityDetailsByJobId(string JobID);

        SAWActivityResponse GetSAWActivityById(string SAWActivityID);
        SAWActivityResponse GetSAWActivityByJobId(string JobID);

        SAWActivityApprovalResponse GetSAWActivityApprovalById(string SAWActivityApprovalID);
        SAWActivityApprovalResponse GetSAWActivityApprovalByJobId(string JobID);
        SAWActivityApprovalResponse GetSAWActivityApprovalByActivityId(string ActivityID);

        ClaimlineResponse GetClaimlineById(string ClaimlineID);
        ClaimlineResponse GetClaimlineByJobId(string JobID);
        ClaimlineResponse GetClaimlineByClaimheaderId(string ClaimheaderID);


        ClaimheaderResponse GetClaimheaderById(string ClaimheaderID);


    }
}
