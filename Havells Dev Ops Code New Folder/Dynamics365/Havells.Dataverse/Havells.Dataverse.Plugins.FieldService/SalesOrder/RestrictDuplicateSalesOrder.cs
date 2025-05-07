using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.SalesOrder
{
    public class RestrictDuplicateSalesOrder : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Setup
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            try
            {
                if (context.MessageName.ToLower() == "create" && context.PrimaryEntityName.ToLower() == "salesorder")
                {
                    Entity _salesOrder = (Entity)context.InputParameters["Target"];
                    EntityReference _custRef = _salesOrder.GetAttributeValue<EntityReference>("customerid");
                    EntityReference _sellingsource = _salesOrder.GetAttributeValue<EntityReference>("hil_sellingsource");
                    EntityReference _orderType = _salesOrder.GetAttributeValue<EntityReference>("hil_ordertype");
                    string _mobNumber = string.Empty;
                    string _fetchXml = string.Empty;
                    string _existingOrderNumber = string.Empty;

                    if (_orderType.Id == new Guid("1F9E3353-0769-EF11-A670-0022486E4ABB") && (_sellingsource.Id == new Guid("03B5A2D6-CC64-ED11-9562-6045BDAC526A") || _sellingsource.Id == new Guid("668E899B-A8A3-ED11-AAD1-6045BDAD27A7"))) //Check if Order Type is AMC Sales Order
                    {
                        Entity _entCont = service.Retrieve(_custRef.LogicalName, _custRef.Id, new ColumnSet("mobilephone"));
                        _mobNumber = _entCont.Contains("mobilephone") ? _entCont.GetAttributeValue<string>("mobilephone") : string.Empty;

                        _fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='salesorder'>
                            <attribute name='name' />
                            <attribute name='salesorderid' />                               
                            <attribute name='createdon' />
                            <order attribute='createdon' descending='true' />
                            <filter type='and'>
                                <condition attribute='hil_ordertype' operator='eq' value='{_orderType.Id}' />
                                <condition attribute='customerid' operator='eq' value='{_entCont.Id}' />
                                <condition attribute='hil_paymentstatus' operator='ne' value='2' />
                                <condition attribute='statecode' operator='ne' value='2' />
                                <condition attribute='hil_sellingsource' operator='in'>
                                    <value>{{03B5A2D6-CC64-ED11-9562-6045BDAC526A}}</value>
                                    <value>{{668E899B-A8A3-ED11-AAD1-6045BDAD27A7}}</value>
                                </condition>
                            </filter>
                            </entity>
                            </fetch>";

                        EntityCollection existingOrders = service.RetrieveMultiple(new FetchExpression(_fetchXml));
                        if (existingOrders.Entities.Count > 0)
                        {
                            _existingOrderNumber = existingOrders.Entities[0].GetAttributeValue<string>("name");

                            throw new InvalidPluginExecutionException($"Duplicate AMC Order found!!! An Open AMC Order exist with Order# {_existingOrderNumber}.");
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
