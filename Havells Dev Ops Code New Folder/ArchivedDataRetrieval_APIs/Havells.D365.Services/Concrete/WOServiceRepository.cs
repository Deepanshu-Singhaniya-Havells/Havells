using Havells.D365.Data;
using Havells.D365.Entities.Common;
using Havells.D365.Entities.WorkOrderService.Entity;
using Havells.D365.Entities.WorkOrderService.Request;
using Havells.D365.Entities.WorkOrderService.Response;
using Havells.D365.Services.Abstract;
using Havells.D365.Services.Utility;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace Havells.D365.Services.Concrete
{
    public class WOServiceRepository : IWOServiceRepository
    {
        private IConfiguration configuration;
        public WOServiceRepository(IConfiguration configuration)
        {
            this.configuration = configuration;
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

    }
}
