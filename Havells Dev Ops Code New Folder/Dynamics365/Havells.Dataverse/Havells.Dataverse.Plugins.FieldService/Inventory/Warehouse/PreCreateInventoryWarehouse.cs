using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreateInventoryWarehouse : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorywarehouse" && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                ProcessRequest(entity, _tracingService, service);
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Guid _accountId = entity.GetAttributeValue<EntityReference>("hil_franchise").Id;

                string _query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='account'>
                    <attribute name='accountnumber' />
                    <attribute name='ownerid' />
                    <filter type='and'>
                        <condition attribute='accountid' operator='eq' value='{_accountId}' />
                    </filter>
                    </entity>
                    </fetch>";

                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_query));
                if (_entCol.Entities.Count > 0) {
                    entity["hil_name"] = $"{_entCol.Entities[0].GetAttributeValue<string>("accountnumber")} - {entity.FormattedValues["hil_type"]}";
                    entity["ownerid"] = _entCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
