using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.WorkOrderIncident
{
    public class PreValidateUpdatePlugin : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "msdyn_workorderincident" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("statecode") && entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        Entity _entWOIncident = service.Retrieve(entity.LogicalName, entity.Id,new ColumnSet("msdyn_workorder"));
                        EntityReference _entRefWO = _entWOIncident.GetAttributeValue<EntityReference>("msdyn_workorder");
                        Entity _entWO = service.Retrieve(_entRefWO.LogicalName, _entRefWO.Id, new ColumnSet("msdyn_substatus"));
                        EntityReference _entRefWOStatus = _entWO.GetAttributeValue<EntityReference>("msdyn_substatus");
                        if (_entRefWOStatus.Name == "Canceled" || _entRefWOStatus.Name == "Closed" || _entRefWOStatus.Name == "Work Done" || _entRefWOStatus.Name == "KKG Audit Failed")
                            throw new Exception("Access denied! You are not allowed to deactivate this Record.");
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
