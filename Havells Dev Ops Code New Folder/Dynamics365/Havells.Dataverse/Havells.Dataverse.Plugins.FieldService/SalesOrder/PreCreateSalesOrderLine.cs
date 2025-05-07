using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.SalesOrder
{
    public class PreCreateSalesOrderLine : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "salesorderdetail")
            {
                tracingService.Trace("Target");
                Entity entity = (Entity)context.InputParameters["Target"];

                if (entity.Contains("salesorderid"))
                {
                    Guid _orderType = new Guid("1F9E3353-0769-EF11-A670-0022486E4ABB");
                    EntityReference entRefOrder = (EntityReference)entity["salesorderid"];

                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='salesorderdetail'>
                        <attribute name='salesorderdetailid' />
                        <filter type='and'>
                            <condition attribute='salesorderid' operator='eq' value='{entRefOrder.Id}' />
                        </filter>
                        <link-entity name='salesorder' from='salesorderid' to='salesorderid' link-type='inner' alias='ac'>
                            <filter type='and'>
                            <condition attribute='hil_ordertype' operator='eq' value='{_orderType}' />
                            </filter>
                        </link-entity>
                        </entity>
                        </fetch>";
                    EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (_entCol.Entities.Count >= 1)
                    {
                        throw new InvalidPluginExecutionException("Multiple AMCs are not allowed in Single Order.");
                    }
                }
            }
        }
    }
}

