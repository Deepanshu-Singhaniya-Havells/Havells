using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreateWorkOrderIncident : IPlugin
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
                Entity entity = (Entity)context.InputParameters["Target"];
                ProcessRequest(entity, _tracingService, service);
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Entity _jobHeader = service.Retrieve("msdyn_workorder", entity.GetAttributeValue<EntityReference>("msdyn_workorder").Id, new ColumnSet("msdyn_name"));

                string _query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorderincident'>
                <attribute name='msdyn_workorderincidentid' />
                <order attribute='msdyn_name' descending='false' />
                <filter type='and'>
                    <condition attribute='msdyn_workorder' operator='eq' value='{_jobHeader.Id}' />
                </filter>
                </entity>
                </fetch>";

                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_query));
                int num = _entCol.Entities.Count + 1;
                entity["msdyn_name"] = $"{_jobHeader.GetAttributeValue<string>("msdyn_name")}-I{num.ToString().PadLeft(3, '0')}";
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
