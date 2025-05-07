using Havells.D365.Data;
using Havells.D365.Entities.Common;
using Havells.D365.Entities.WorkorderProduct.Entity;
using Havells.D365.Entities.WorkorderProduct.Request;
using Havells.D365.Entities.WorkorderProduct.Response;
using Havells.D365.Services.Abstract;
using Havells.D365.Services.Utility;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace Havells.D365.Services.Concrete
{
    public class WorkOrderProductRepository : IWorkOrderProductRepository
    {
        private IConfiguration configuration;
        public WorkOrderProductRepository(IConfiguration configuration)
        {
            this.configuration = configuration;
        }


        //usp_GetWOProductDetailByID    WorkOrderProductID
        //usp_GetWOProductDetailByIncidentID    WorkOrderIncidentID
        //usp_GetWOProductDetailByJobID WorkOrderID
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
    }
}
