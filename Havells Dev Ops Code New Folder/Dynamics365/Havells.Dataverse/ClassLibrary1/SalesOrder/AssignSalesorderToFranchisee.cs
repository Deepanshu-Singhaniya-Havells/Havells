using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.CustomConnector.SalesOrder
{
    public class AssignSalesorderToFranchisee : IPlugin
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
            string _message = "Assignment Matrix does not exist.";
            string _status = "False";
            if (context.InputParameters.Contains("OrderId"))
            {
                try
                {
                    tracingService.Trace("Execution Start");
                    string OrderId = context.InputParameters["OrderId"].ToString();
                    tracingService.Trace("Order Record ID " + OrderId);
                    if (string.IsNullOrWhiteSpace(OrderId))
                    {
                        _message = "Invalid OrderId";
                    }
                    else
                    {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='salesorder'>
                        <attribute name='salesorderid' />
                        <attribute name='hil_productdivision' />
                        <filter type='and'>
                        <condition attribute='name' operator='eq' value='{OrderId}' />
                        </filter>
                        <link-entity name='hil_address' from='hil_addressid' to='hil_serviceaddress' visible='false' link-type='outer' alias='ad'>
                        <attribute name='hil_pincode' />
                        </link-entity>
                        </entity>
                        </fetch>";
                        EntityCollection ecoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (ecoll.Entities.Count > 0)
                        {
                            EntityReference productdivision = ecoll.Entities[0].GetAttributeValue<EntityReference>("hil_productdivision");
                            EntityReference pincode = (EntityReference)ecoll.Entities[0].GetAttributeValue<AliasedValue>("ad.hil_pincode").Value;
                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                            <entity name='hil_assignmentmatrix'>
                            <attribute name='hil_franchiseedirectengineer' />
                            <attribute name='hil_assignmentmatrixid' />
                            <order attribute='hil_assignmentmatrixid' descending='true' />
                            <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_callsubtypename' operator='like' value='%AMC%' />
                            <condition attribute='hil_division' operator='eq' value='{productdivision.Id}' />
                            <condition attribute='hil_pincode' operator='eq' value='{pincode.Id}' />
                            </filter>
                            <link-entity name='account' from='accountid' to='hil_franchiseedirectengineer' link-type='inner' alias='acc'>
                            <attribute name='ownerid'/>
                             </link-entity>
                            </entity>
                            </fetch>";
                            EntityCollection ecoll1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (ecoll1.Entities.Count > 0)
                            {
                                if (ecoll1.Entities[0].Contains("hil_franchiseedirectengineer"))
                                {
                                    EntityReference _owner = ecoll1.Entities[0].GetAttributeValue<EntityReference>("hil_franchiseedirectengineer");

                                    Entity _entOrder = new Entity(ecoll.Entities[0].LogicalName, ecoll.Entities[0].Id);
                                    _entOrder["ownerid"] = (EntityReference)ecoll1.Entities[0].GetAttributeValue<AliasedValue>("acc.ownerid").Value;//_owner;
                                    service.Update(_entOrder);
                                    _message = "Sucess";
                                    _status = "True";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _message = ex.Message;
                }
                context.OutputParameters["Message"] = _message;
                context.OutputParameters["Status"] = _status;
            }
        }
    }
}

