using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.AMC_OmniChannel
{
    public class InvoicePostCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    tracingService.Trace("11");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_productcode", "pricelevelid", "transactioncurrencyid", "ownerid", "name"));
                    tracingService.Trace("14");
                    Entity productEnt = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_productcode").LogicalName,
                        entity.GetAttributeValue<EntityReference>("hil_productcode").Id,
                        new ColumnSet("description", "defaultuomid"));
                    tracingService.Trace("13");
                    Entity _entInvoiceLine = new Entity("invoicedetail");
                    _entInvoiceLine["quantity"] = new decimal(1);
                    tracingService.Trace("1");
                    _entInvoiceLine["salesrepid"] = entity.GetAttributeValue<EntityReference>("ownerid");
                    _entInvoiceLine["invoiceid"] = entity.ToEntityReference();
                    _entInvoiceLine["productid"] = entity.GetAttributeValue<EntityReference>("hil_productcode");
                    _entInvoiceLine["uomid"] = productEnt.GetAttributeValue<EntityReference>("defaultuomid");
                    _entInvoiceLine["transactioncurrencyid"] = entity.GetAttributeValue<EntityReference>("transactioncurrencyid");
                    _entInvoiceLine["msdyn_linetype"] = new OptionSetValue(690970001);
                    _entInvoiceLine["description"] = productEnt.GetAttributeValue<string>("description");
                    _entInvoiceLine["msdynce_invoicenumber"] = entity.GetAttributeValue<string>("name");
                    _entInvoiceLine["msdyn_lineorder"] = 1;
                    service.Create(_entInvoiceLine);
                    tracingService.Trace("12");
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("InvoiceUpBilling.UpBilling.Invoice.Execute Error " + ex.Message);
            }

        }

    }
}
