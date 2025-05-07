using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidationPurchaseReceiptLine : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderreceiptline" && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                ProcessRequest(entity, _tracingService, service);
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                Entity _headerId = service.Retrieve("hil_inventorypurchaseorderreceipt", entity.GetAttributeValue<EntityReference>("hil_receiptnumber").Id, new ColumnSet("hil_receiptstatus"));
                if (entity.Contains("hil_receiptstatus"))
                {
                    OptionSetValue _adjStatus = entity.GetAttributeValue<OptionSetValue>("hil_receiptstatus");
                    if (_adjStatus.Value == 2)//Posted
                    {
                        throw new InvalidPluginExecutionException("Aceed denied. Receipt is already Posted.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
