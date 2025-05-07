using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.AMC_OmniChannel
{
    public class InvoicePreCreate : IPlugin
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
                    if (entity.Contains("customerid"))
                    {
                        EntityReference customer = entity.GetAttributeValue<EntityReference>("customerid");
                        if (customer.LogicalName == "contact")
                        {
                            Entity contact = service.Retrieve(customer.LogicalName, customer.Id, new ColumnSet("mobilephone"));
                            entity["hil_mobilenumber"] =contact.GetAttributeValue<string>("mobilephone");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPLugin.AMC_OmniChannel.Invoice.Execute Error " + ex.Message);
            }

        }

    }
}
