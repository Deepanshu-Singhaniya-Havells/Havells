using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidateWorkOrderIncident : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion


            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "msdyn_workorderincident" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (!entity.Contains("msdyn_customerasset"))
                        throw new InvalidPluginExecutionException("Customer Asset is required.");

                    if (!entity.Contains("hil_natureofcomplaint"))
                        throw new InvalidPluginExecutionException("NOC is required.");

                    if (!entity.Contains("hil_observation"))
                        throw new InvalidPluginExecutionException("Observation is required.");

                    if (!entity.Contains("msdyn_incidenttype"))
                        throw new InvalidPluginExecutionException("Cause is required.");

                    if (!entity.Contains("msdyn_workorder"))
                        throw new InvalidPluginExecutionException("Job Id is required.");

                    Entity entityHeader = service.Retrieve("msdyn_workorder", entity.GetAttributeValue<EntityReference>("msdyn_workorder").Id, new ColumnSet("msdyn_substatus"));
                    EntityReference _jobSubstatus = entityHeader.GetAttributeValue<EntityReference>("msdyn_substatus");

                    if (_jobSubstatus.Id == new Guid("1727fa6c-fa0f-e911-a94e-000d3af060a1") || _jobSubstatus.Id == new Guid("1527fa6c-fa0f-e911-a94e-000d3af060a1") || _jobSubstatus.Id == new Guid("6c8f2123-5106-ea11-a811-000d3af057dd") || _jobSubstatus.Id == new Guid("2927fa6c-fa0f-e911-a94e-000d3af060a1"))
                    {  //Emergency Order and Approved
                        throw new Exception("Access Denied!!!.\nYou can't add Incident in Completed Job.");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
