using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreatePurchaseOrderLine : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderline" && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                ProcessRequest(entity, _tracingService, service);
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Entity _poHeader = service.Retrieve("hil_inventorypurchaseorder", entity.GetAttributeValue<EntityReference>("hil_ponumber").Id, new ColumnSet("hil_name"));

                string _query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_inventorypurchaseorderline'>
                <attribute name='hil_inventorypurchaseorderlineid' />
                <order attribute='hil_name' descending='false' />
                <filter type='and'>
                    <condition attribute='hil_ponumber' operator='eq' value='{_poHeader.Id}' />
                </filter>
                </entity>
                </fetch>";

                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_query));
                int num = _entCol.Entities.Count + 1;
                entity["hil_name"] = $"{_poHeader.GetAttributeValue<string>("hil_name")}-{num.ToString().PadLeft(3, '0')}";
                entity["hil_partstatus"] = new OptionSetValue(1);//Requested
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
