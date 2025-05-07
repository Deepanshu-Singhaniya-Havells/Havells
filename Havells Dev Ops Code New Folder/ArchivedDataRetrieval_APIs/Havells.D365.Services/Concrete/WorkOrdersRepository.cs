using Havells.D365.Data;
using Havells.D365.Entities.Common;
using Havells.D365.Entities.WorkOrders.Entity;
using Havells.D365.Entities.WorkOrders.Request;
using Havells.D365.Entities.WorkOrders.Response;
using Havells.D365.Entities.Incident.Entity;
using Havells.D365.Entities.Incident.Response;
using Havells.D365.Entities.WorkorderProduct.Entity;
using Havells.D365.Entities.WorkorderProduct.Response;
using Havells.D365.Entities.WorkOrderService.Entity;
using Havells.D365.Entities.WorkOrderService.Response;
using Havells.D365.Services.Abstract;
using Havells.D365.Services.Utility;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using Havells.D365.Entities.POStatusTracker.Response;
using Havells.D365.Entities.POStatusTracker.Entity;
using Havells.D365.Entities.ProductRequest.Response;
using Havells.D365.Entities.ProductRequest.Entity;
using Havells.D365.Entities.ProductReqHeader.Response;
using Havells.D365.Entities.ProductReqHeader.Entity;
using Havells.D365.Entities.ActivityDetails.Response;
using Havells.D365.Entities.ActivityDetails.Entity;

namespace Havells.D365.Services.Concrete
{
    public class WorkOrdersRepository : IWorkOrderRepository
    {
        private IConfiguration configuration;
        public WorkOrdersRepository(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        public WorkOrderResponse GetWorkOrderByCustRef(string custRef, int pageNumber, int recordsPerPage)
        {
            WorkOrderResponse response = new WorkOrderResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@PageNumber",pageNumber),
                    new SqlParameter("@RecordsPerPage",recordsPerPage),
                    new SqlParameter("@CustmerRef",custRef)
                };
                var result = Utilities.ConvertDataTable<dtoWorkOrders>(SqlHelper.ExecuteProc(configuration.GetConnectionString("ConStr"), CommonObject.usp_GetWorkOrderDetailByCustomerRef, parameter).Tables[0]);
                response.workOrders = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.workOrders = new List<dtoWorkOrders>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public WorkOrderResponse GetWorkOrderById(string workOrderId)
        {
            WorkOrderResponse response = new WorkOrderResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@WorkOrderId",workOrderId)
                };

                var result = Utilities.ConvertDataTable<dtoWorkOrders>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetWorkOrderDetailByWorkOrderId, parameter).
                    Tables[0]);
                response.workOrders = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.workOrders = new List<dtoWorkOrders>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public WorkOrderResponse GetWorkOrderByName(string workOrderName)
        {
            WorkOrderResponse response = new WorkOrderResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@Name",workOrderName)
                };
                var result = Utilities.ConvertDataTable<dtoWorkOrders>(SqlHelper.ExecuteProc(configuration.GetConnectionString("ConStr"), CommonObject.usp_GetWorkOrderDetailByName, parameter).Tables[0]);
                response.workOrders = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.workOrders = new List<dtoWorkOrders>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public IncidentResponse GetIncidentById(string incidentId)
        {
            IncidentResponse response = new IncidentResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@IncidentID",incidentId)
                };

                var result = Utilities.ConvertDataTable<dtoIncidents>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetIncidentDetailByID, parameter).
                    Tables[0]);
                response.incident = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.incident = new List<dtoIncidents>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public IncidentResponse GetIncidentByWorkOrderID(string workorderId)
        {
            IncidentResponse response = new IncidentResponse();
            try
            {
                SqlParameter[] parameter =
                {
                   new SqlParameter("WorkOrderId",workorderId)
                };
                var result = Utilities.ConvertDataTable<dtoIncidents>(SqlHelper.ExecuteProc(
                    configuration.GetConnectionString("ConStr"),
                    CommonObject.usp_GetIncidentDetailByWorkorderID,
                    parameter).Tables[0]);
                response.incident = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.incident = new List<dtoIncidents>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public WorkOrderProductResponse GetWOProductDetailByID(string WorkOrderProductID)
        {
            WorkOrderProductResponse response = new WorkOrderProductResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@WorkOrderProductID",WorkOrderProductID)
                };

                var result = Utilities.ConvertDataTable<dtoWorkorderProduct>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetWOProductDetailByID, parameter).
                    Tables[0]);
                response.incident = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.incident = new List<dtoWorkorderProduct>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public WorkOrderProductResponse GetWOProductDetailByIncidentID(string WorkOrderIncidentID)
        {
            WorkOrderProductResponse response = new WorkOrderProductResponse();
            try
            {
                SqlParameter[] parameter =
                {
                   new SqlParameter("@WorkOrderIncidentID",WorkOrderIncidentID)
                };
                var result = Utilities.ConvertDataTable<dtoWorkorderProduct>(SqlHelper.ExecuteProc(
                    configuration.GetConnectionString("ConStr"),
                    CommonObject.usp_GetWOProductDetailByIncidentID,
                    parameter).Tables[0]);
                response.incident = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.incident = new List<dtoWorkorderProduct>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public WorkOrderProductResponse GetWOProductDetailByJobID(string WorkOrderID)
        {
            WorkOrderProductResponse response = new WorkOrderProductResponse();
            try
            {
                SqlParameter[] parameter =
                {
                   new SqlParameter("WorkOrderID",WorkOrderID)
                };
                var result = Utilities.ConvertDataTable<dtoWorkorderProduct>(SqlHelper.ExecuteProc(
                    configuration.GetConnectionString("ConStr"),
                    CommonObject.usp_GetWOProductDetailByJobID,
                    parameter).Tables[0]);
                response.incident = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.incident = new List<dtoWorkorderProduct>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public WorkOrderServiceResponse GetWOServiceDetailByJobID(string WorkOrderID)
        {
            WorkOrderServiceResponse response = new WorkOrderServiceResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@WorkOrderID",WorkOrderID)
                };

                var result = Utilities.ConvertDataTable<dtoWorkOrderService>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetWOServiceDetailByJobID, parameter).
                    Tables[0]);
                response.OrderService = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.OrderService = new List<dtoWorkOrderService>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public WorkOrderServiceResponse GetWOServiceDetailByID(string WorkOrderServiceID)
        {
            WorkOrderServiceResponse response = new WorkOrderServiceResponse();
            try
            {
                SqlParameter[] parameter =
                {
                   new SqlParameter("WorkOrderServiceID",WorkOrderServiceID)
                };
                var result = Utilities.ConvertDataTable<dtoWorkOrderService>(SqlHelper.ExecuteProc(
                    configuration.GetConnectionString("ConStr"),
                    CommonObject.usp_GetWOServiceDetailByID,
                    parameter).Tables[0]);
                response.OrderService = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.OrderService = new List<dtoWorkOrderService>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public WorkOrderServiceResponse GetWOServiceDetailByIncidentID(string WorkOrderIncidentID)
        {
            WorkOrderServiceResponse response = new WorkOrderServiceResponse();
            try
            {
                SqlParameter[] parameter =
                {
                   new SqlParameter("WorkOrderIncidentID",WorkOrderIncidentID)
                };
                var result = Utilities.ConvertDataTable<dtoWorkOrderService>(SqlHelper.ExecuteProc(
                    configuration.GetConnectionString("ConStr"),
                    CommonObject.usp_GetWOServiceDetailByIncidentID,
                    parameter).Tables[0]);
                response.OrderService = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.OrderService = new List<dtoWorkOrderService>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public JobSheetResponse GetJobSheetById(string JobId)
        {
            JobSheetResponse response = new JobSheetResponse();
            List<DtoJobSheetData> jobResponse = new List<DtoJobSheetData>();
            DtoJobSheetData _jobResponse = new DtoJobSheetData();

            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@JOBID",JobId)
                };
                var jobData = Utilities.ConvertDataTable<dtoJob_JobSheetData>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.USP_JOBDATA_GETJOBSHEET, new SqlParameter[] { new SqlParameter("@JOBID", JobId) }).
                    Tables[0]);
                var incidentdata = Utilities.ConvertDataTable<dtoJobIncident_JobSheetData>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.USP_JOBINCIDENTDATA_GETJOBSHEET, new SqlParameter[] { new SqlParameter("@JOBID", JobId) }).
                    Tables[0]);
                var servidedata = Utilities.ConvertDataTable<DtoJobSparePart_JobSheetData>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.USP_JOBSERVICEANDSPAREDATA_GETJOBSHEET, new SqlParameter[] { new SqlParameter("@JOBID", JobId) }).
                    Tables[0]);

                _jobResponse.Job = jobData;
                _jobResponse.IncidentData = incidentdata;
                _jobResponse.SparePartAndProductData = servidedata;
                jobResponse.Add(_jobResponse);
                response.workOrders = jobResponse;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.workOrders = new List<DtoJobSheetData>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public JobSheetCountResponse GetJobCountByCustomerAssetID(string CustomerAssetID)
        {
            JobSheetCountResponse response = new JobSheetCountResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@CUSTOMERASSETID",CustomerAssetID)
                };

                var result = Utilities.ConvertDataTable<dtoJobCount>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.USP_GETJOBCOUNT, parameter).
                    Tables[0]);
                response.customerasset = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.customerasset = new List<dtoJobCount>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public POStatusTrackerResponse GetPOStatusTrackerById(string POStatusTrackerID)
        {
            POStatusTrackerResponse response = new POStatusTrackerResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@postatustrackerid",POStatusTrackerID)
                };

                var result = Utilities.ConvertDataTable<dtoPOStatusTracker>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetPOStatusTrackerByID, parameter).
                    Tables[0]);
                response.POStatusTracker = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.POStatusTracker = new List<dtoPOStatusTracker>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public POStatusTrackerResponse GetPOStatusTrackerByJobId(string JobID)
        {
            POStatusTrackerResponse response = new POStatusTrackerResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@jobid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoPOStatusTracker>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetPOStatusTrackerByJobID, parameter).
                    Tables[0]);
                response.POStatusTracker = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.POStatusTracker = new List<dtoPOStatusTracker>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public POStatusTrackerResponse GetPOStatusTrackerByPRHeader(string PRHeader)
        {
            POStatusTrackerResponse response = new POStatusTrackerResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@prheaderno",PRHeader)
                };

                var result = Utilities.ConvertDataTable<dtoPOStatusTracker>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetPOStatusTrackerByPRHeader, parameter).
                    Tables[0]);
                response.POStatusTracker = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.POStatusTracker = new List<dtoPOStatusTracker>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }


        public ProductRequestResponse GetProductRequestById(string POStatusTrackerID)
        {
            ProductRequestResponse response = new ProductRequestResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@productrequestid",POStatusTrackerID)
                };

                var result = Utilities.ConvertDataTable<dtoProductRequest>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetProductRequestByID, parameter).
                    Tables[0]);
                response.ProductRequest = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ProductRequest = new List<dtoProductRequest>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public ProductRequestResponse GetProductRequestByJobId(string JobID)
        {
            ProductRequestResponse response = new ProductRequestResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@jobid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoProductRequest>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetProductRequestByJobID, parameter).
                    Tables[0]);
                response.ProductRequest = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ProductRequest = new List<dtoProductRequest>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public ProductRequestResponse GetProductRequestByPRHeader(string PRHeader)
        {
            ProductRequestResponse response = new ProductRequestResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@prheaderno",PRHeader)
                };

                var result = Utilities.ConvertDataTable<dtoProductRequest>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetProductRequestByPRHeader, parameter).
                    Tables[0]);
                response.ProductRequest = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ProductRequest = new List<dtoProductRequest>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public ProductReqHeaderResponse GetProductReqHeaderById(string ProductReqHeaderID)
        {
            ProductReqHeaderResponse response = new ProductReqHeaderResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@productreqheaderid",ProductReqHeaderID)
                };

                var result = Utilities.ConvertDataTable<dtoProductReqHeader>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetProductReqHeaderByID, parameter).
                    Tables[0]);
                response.ProductReqHeader = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ProductReqHeader = new List<dtoProductReqHeader>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public ProductReqHeaderResponse GetProductReqHeaderByJobId(string JobID)
        {
            ProductReqHeaderResponse response = new ProductReqHeaderResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@jobid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoProductReqHeader>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetProductReqHeaderByJobID, parameter).
                    Tables[0]);
                response.ProductReqHeader = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ProductReqHeader = new List<dtoProductReqHeader>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public ActivityDetailsResponse GetActivityDetailsById(string ActivityID)
        {
            ActivityDetailsResponse response = new ActivityDetailsResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@activityid",ActivityID)
                };

                var result = Utilities.ConvertDataTable<dtoActivityDetails>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetActivityDetailsById, parameter).
                    Tables[0]);
                response.ActivityDetails = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ActivityDetails = new List<dtoActivityDetails>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public ActivityDetailsResponse GetActivityDetailsByJobId(string JobID)
        {
            ActivityDetailsResponse response = new ActivityDetailsResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@regardingobjectid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoActivityDetails>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetActivityDetailsByJobId, parameter).
                    Tables[0]);
                response.ActivityDetails = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.ActivityDetails = new List<dtoActivityDetails>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public SAWActivityResponse GetSAWActivityById(string SAWActivityID)
        {
            SAWActivityResponse response = new SAWActivityResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@sawactivityid",SAWActivityID)
                };

                var result = Utilities.ConvertDataTable<dtoSAWActivity>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetSAWActivityById, parameter).
                    Tables[0]);
                response.SAWActivity = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.SAWActivity = new List<dtoSAWActivity>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public SAWActivityResponse GetSAWActivityByJobId(string JobID)
        {
            SAWActivityResponse response = new SAWActivityResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@jobid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoSAWActivity>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetSAWActivityByJobId, parameter).
                    Tables[0]);
                response.SAWActivity = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.SAWActivity = new List<dtoSAWActivity>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public SAWActivityApprovalResponse GetSAWActivityApprovalById(string SAWActivityApprovalID)
        {
            SAWActivityApprovalResponse response = new SAWActivityApprovalResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@sawactivityapprovalid",SAWActivityApprovalID)
                };

                var result = Utilities.ConvertDataTable<dtoSAWActivityApproval>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetSAWActivityApprovalById, parameter).
                    Tables[0]);
                response.SAWActivityApproval = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.SAWActivityApproval = new List<dtoSAWActivityApproval>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public SAWActivityApprovalResponse GetSAWActivityApprovalByJobId(string JobID)
        {
            SAWActivityApprovalResponse response = new SAWActivityApprovalResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@jobid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoSAWActivityApproval>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetSAWActivityApprovalByJobId, parameter).
                    Tables[0]);
                response.SAWActivityApproval = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.SAWActivityApproval = new List<dtoSAWActivityApproval>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public SAWActivityApprovalResponse GetSAWActivityApprovalByActivityId(string ActivityID)
        {
            SAWActivityApprovalResponse response = new SAWActivityApprovalResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@sawactivityId",ActivityID)
                };

                var result = Utilities.ConvertDataTable<dtoSAWActivityApproval>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetSAWActivityApprovalByActivityId, parameter).
                    Tables[0]);
                response.SAWActivityApproval = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.SAWActivityApproval = new List<dtoSAWActivityApproval>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public ClaimlineResponse GetClaimlineById(string ClaimlineID)
        {
            ClaimlineResponse response = new ClaimlineResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@sawactivityid",ClaimlineID)
                };

                var result = Utilities.ConvertDataTable<dtoClaimline>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetClaimlineById, parameter).
                    Tables[0]);
                response.Claimline = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.Claimline = new List<dtoClaimline>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public ClaimlineResponse GetClaimlineByJobId(string JobID)
        {
            ClaimlineResponse response = new ClaimlineResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@jobid",JobID)
                };

                var result = Utilities.ConvertDataTable<dtoClaimline>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetClaimlineByJobId, parameter).
                    Tables[0]);
                response.Claimline = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.Claimline = new List<dtoClaimline>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }
        public ClaimlineResponse GetClaimlineByClaimheaderId(string ClaimheaderID)
        {
            ClaimlineResponse response = new ClaimlineResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@claimheader",ClaimheaderID)
                };

                var result = Utilities.ConvertDataTable<dtoClaimline>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetClaimlineByClaimheaderId, parameter).
                    Tables[0]);
                response.Claimline = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.Claimline = new List<dtoClaimline>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }

        public ClaimheaderResponse GetClaimheaderById(string ClaimheaderID)
        {
            ClaimheaderResponse response = new ClaimheaderResponse();
            try
            {
                SqlParameter[] parameter =
                {
                    new SqlParameter("@sawactivityapprovalid",ClaimheaderID)
                };

                var result = Utilities.ConvertDataTable<dtoClaimheader>(SqlHelper.ExecuteProc
                    (configuration.GetConnectionString("ConStr"), CommonObject.usp_GetClaimheaderById, parameter).
                    Tables[0]);
                response.Claimheader = result;
                response.Success = true;
            }
            catch (Exception ex)
            {
                response.Claimheader = new List<dtoClaimheader>();
                response.Success = false;
                response.Error = ex.Message.ToString();
            }
            return response;
        }


    }
}
