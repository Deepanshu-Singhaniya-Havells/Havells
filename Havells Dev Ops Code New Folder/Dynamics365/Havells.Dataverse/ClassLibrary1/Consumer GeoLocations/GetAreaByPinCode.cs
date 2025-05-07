using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class GetAreaByPinCode : IPlugin
    {
        IoTAreas objArea = null;
        List<IoTAreas> lstAreas = null;
        public void Execute(IServiceProvider serviceProvider)

        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string JsonResponse = "";
            Guid PINCodeGuid;
            List<IoTAreas> IoTAreas = new List<IoTAreas>();
            if (context.InputParameters.Contains("PINCodeGuid"))
            {
                string pincodeGuidStr = Convert.ToString(context.InputParameters["PINCodeGuid"]);
                // Validate if PINCodeGuid is a valid GUID
                if (!Guid.TryParse(pincodeGuidStr, out PINCodeGuid))
                {
                    objArea = new IoTAreas { StatusCode = "204", StatusDescription = "Invalid or missing Pin Code Guid." };
                    JsonResponse = JsonSerializer.Serialize(new List<IoTAreas> { objArea });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                JsonResponse = JsonSerializer.Serialize(IoTGetAreasByPinCode(PINCodeGuid, service));
                context.OutputParameters["data"] = JsonResponse;
            }
        }
        public List<IoTAreas> IoTGetAreasByPinCode(Guid PINCodeGuid, IOrganizationService service)
        {
            try
            {
                if (service != null)
                {
                    QueryExpression query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_pincode", "hil_area");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, PINCodeGuid);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active records
                    EntityCollection entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objArea = new IoTAreas { StatusCode = "204", StatusDescription = "No Area found." };
                        lstAreas = new List<IoTAreas> { objArea };
                    }
                    else
                    {
                        lstAreas = new List<IoTAreas>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            if (ent.Attributes.Contains("hil_area"))
                            {
                                lstAreas.Add(new IoTAreas
                                {
                                    PINCodeGuid = PINCodeGuid,
                                    AreaName = ent.GetAttributeValue<EntityReference>("hil_area").Name,
                                    AreaGuid = ent.GetAttributeValue<EntityReference>("hil_area").Id,
                                    StatusCode = "200",
                                    StatusDescription = "OK"
                                });
                            }
                        }
                    }
                }
                else
                {
                    objArea = new IoTAreas { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstAreas = new List<IoTAreas> { objArea };
                }
            }
            catch (Exception ex)
            {
                objArea = new IoTAreas { StatusCode = "500", StatusDescription = "D365 Internal Server Error: " + ex.Message };
                lstAreas = new List<IoTAreas> { objArea };
            }
            return lstAreas;
        }
        public class IoTAreas
        {
            public Guid PINCodeGuid { get; set; }
            public Guid AreaGuid { get; set; }
            public string AreaName { get; set; }
            public string StatusCode { get; set; }
            public string StatusDescription { get; set; }
        }
    }
}
