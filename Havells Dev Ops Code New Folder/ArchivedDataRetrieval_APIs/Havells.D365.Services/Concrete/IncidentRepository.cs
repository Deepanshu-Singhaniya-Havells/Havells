using Havells.D365.Data;
using Havells.D365.Entities.Common;
using Havells.D365.Entities.Incident.Entity;
using Havells.D365.Entities.Incident.Request;
using Havells.D365.Entities.Incident.Response;
using Havells.D365.Services.Abstract;
using Havells.D365.Services.Utility;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace Havells.D365.Services.Concrete
{
    public class IncidentRepository : IIncidentRepository
    {
        private IConfiguration configuration;
        public IncidentRepository(IConfiguration configuration)
        {
            this.configuration = configuration;
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

    }
}
