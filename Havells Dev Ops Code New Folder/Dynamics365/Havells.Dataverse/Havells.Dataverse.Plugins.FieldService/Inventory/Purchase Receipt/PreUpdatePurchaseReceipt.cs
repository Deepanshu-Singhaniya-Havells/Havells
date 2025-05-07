using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdatePurchaseReceipt : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderreceipt" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_receiptstatus"))
                    {
                        OptionSetValue _adjStatus = entity.GetAttributeValue<OptionSetValue>("hil_receiptstatus");
                        if (_adjStatus.Value == 2)//Posted
                        {
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_inventorypurchaseorderreceiptline'>
                            <attribute name='hil_inventorypurchaseorderreceiptlineid'/>
                            <order attribute='hil_name' descending='false'/>
                            <filter type='and'>
                                <filter type='or'>
                                    <condition attribute='hil_receiptquantity' operator='null' />
                                    <condition attribute='hil_receiptquantity' operator='eq' value='0' />
                                </filter>
                                <filter type='or'>
                                    <condition attribute='hil_sortquantity' operator='null' />
                                    <condition attribute='hil_sortquantity' operator='eq' value='0' />
                                </filter>
                                <condition attribute='hil_receiptnumber' operator='eq' value='{entity.Id}'/>
                            </filter>
                            </entity>
                            </fetch>";
                            EntityCollection entColChecks = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entColChecks.Entities.Count > 0) {
                                throw new InvalidPluginExecutionException("Please Input Fresh/Defective/Short Quantity in Receipt line.");
                            }
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
}
