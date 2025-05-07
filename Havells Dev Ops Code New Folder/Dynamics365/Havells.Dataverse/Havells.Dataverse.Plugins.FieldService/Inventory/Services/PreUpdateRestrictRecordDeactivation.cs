using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdateRestrictRecordDeactivation : IPlugin
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
                && (context.MessageName.ToUpper() == "UPDATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Contains("statecode") && entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                {
                    if (context.PrimaryEntityName.ToLower() == "msdyn_workorderincident")
                    {
                        EntityReference _entRefJob = entity.GetAttributeValue<EntityReference>("msdyn_workorder");
                        Entity _entJob = service.Retrieve(_entRefJob.LogicalName, _entRefJob.Id, new ColumnSet("msdyn_substatus"));
                        if (_entJob.Contains("msdyn_substatus"))
                        {
                            EntityReference _entRefSubstatus = _entJob.GetAttributeValue<EntityReference>("msdyn_substatus");
                            if (_entRefSubstatus.Name == "Closed" || _entRefSubstatus.Name == "Work Done")
                                throw new InvalidPluginExecutionException("Access denied! You are not allowed to deactivate this Record.");
                        }
                    }
                    else
                        throw new InvalidPluginExecutionException("Access denied! You are not allowed to deactivate this Record.");
                }
            }
        }
    }
}
