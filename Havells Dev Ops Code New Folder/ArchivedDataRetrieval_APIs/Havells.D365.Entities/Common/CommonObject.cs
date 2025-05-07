using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.Common
{
    public class CommonObject
    {
        public const string usp_GetWorkOrderDetailByName = "[D365].[usp_GetWorkOrderDetailByName]";
        public const string usp_GetWorkOrderDetailByWorkOrderId = "[D365].[usp_GetWorkOrderDetailByWorkOrderId]";
        public const string usp_GetWorkOrderDetailByCustomerRef = "[D365].[usp_GetWorkOrderDetailByCustomerRef]";
        public const string usp_GetIncidentDetailByID = "[D365].[usp_GetIncidentDetailByID]";
        public const string usp_GetIncidentDetailByWorkorderID = "[D365].[usp_GetIncidentDetailByWorkorderID]";
       
        public const string usp_GetWOProductDetailByID = "[D365].[usp_GetWOProductDetailByID]";
        public const string usp_GetWOProductDetailByIncidentID = "[D365].[usp_GetWOProductDetailByIncidentID]";
        public const string usp_GetWOProductDetailByJobID = "[D365].[usp_GetWOProductDetailByJobID]";

        public const string usp_GetWOServiceDetailByJobID = "[D365].[usp_GetWOServiceDetailByJobID]";
        public const string usp_GetWOServiceDetailByID = "[D365].[usp_GetWOServiceDetailByID]";
        public const string usp_GetWOServiceDetailByIncidentID = "[D365].[usp_GetWOServiceDetailByIncidentID]";

        public const string USP_GETJOBSHEETDATA = "[D365].[USP_GETJOBSHEETDATA]";
        public const string USP_GETJOBCOUNT = "[D365].[USP_GETJOBCOUNT]";

        public const string USP_JOBDATA_GETJOBSHEET = "[D365].[USP_JOBDATA_GETJOBSHEET]";
        public const string USP_JOBINCIDENTDATA_GETJOBSHEET = "[D365].[USP_JOBINCIDENTDATA_GETJOBSHEET]";
        public const string USP_JOBSERVICEANDSPAREDATA_GETJOBSHEET = "[D365].[USP_JOBSERVICEANDSPAREDATA_GETJOBSHEET]";

        public const string usp_GetPOStatusTrackerByID = "[D365].[usp_GetPOStatusTrackerById]";
        public const string usp_GetPOStatusTrackerByJobID = "[D365].[usp_GetPOStatusTrackerByJobId]";
        public const string usp_GetPOStatusTrackerByPRHeader = "[D365].[usp_GetPOStatusTrackerByPRHeader]";

        public const string usp_GetProductRequestByID = "[D365].[usp_GetProductRequestById]";
        public const string usp_GetProductRequestByJobID = "[D365].[usp_GetProductRequestByJobId]";
        public const string usp_GetProductRequestByPRHeader = "[D365].[usp_GetProductRequestByPRHeader]";

        public const string usp_GetProductReqHeaderByID = "[D365].[usp_GetProductReqHeaderById]";
        public const string usp_GetProductReqHeaderByJobID = "[D365].[usp_GetProductReqHeaderByJobId]";

        public const string usp_GetActivityDetailsById = "[D365].[usp_GetActivityDetailsById]";
        public const string usp_GetActivityDetailsByJobId = "[D365].[usp_GetActivityDetailsByJobId]";

        public const string usp_GetSAWActivityById = "[D365].[usp_GetSAWActivityById]";
        public const string usp_GetSAWActivityByJobId = "[D365].[usp_GetSAWActivityByJobId]";

        public const string usp_GetSAWActivityApprovalById = "[D365].[usp_GetSAWActivityApprovalById]";
        public const string usp_GetSAWActivityApprovalByJobId = "[D365].[usp_GetSAWActivityApprovalByJobId]";
        public const string usp_GetSAWActivityApprovalByActivityId = "[D365].[usp_GetSAWActivityApprovalByActivityId]";

        public const string usp_GetClaimlineById = "[D365].[usp_GetClaimlineById]";
        public const string usp_GetClaimlineByJobId = "[D365].[usp_GetClaimlineByJobId]";
        public const string usp_GetClaimlineByClaimheaderId = "[D365].[usp_GetClaimlineByClaimheaderId]";

        public const string usp_GetClaimheaderById = "[D365].[usp_GetClaimheaderById]";

    }

    public class EndPoints
    {
        public const string GetConsumerSurvey = "GetConsumerSurvey";
        public const string UpdateConsumerSurvey = "UpdateConsumerSurvey";
    }
}
