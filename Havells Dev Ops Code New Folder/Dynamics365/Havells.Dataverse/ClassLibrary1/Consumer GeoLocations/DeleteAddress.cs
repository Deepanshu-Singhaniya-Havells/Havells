using Havells.Dataverse.CustomConnector.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class DeleteAddress : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            Guid AddressGuid = Guid.Empty;
            string JsonResponse = "";
            if (context.InputParameters.Contains("AddressGuid") && context.InputParameters["AddressGuid"] is string)
            {
                bool isValidGuid = Guid.TryParse(Convert.ToString(context.InputParameters["AddressGuid"]), out AddressGuid);
                if (!isValidGuid)
                {
                    JsonResponse = JsonSerializer.Serialize(new IoTAddressBookResultV1
                    {
                        StatusCode = 204,
                        StatusDescription = "Invalid Address Guid.",
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                JsonResponse = JsonSerializer.Serialize(DeleteConsumerAddress(service, AddressGuid));
                _tracingService.Trace(JsonResponse);
                context.OutputParameters["data"] = JsonResponse;
            }
        }
        public IoTAddressBookResultV1 DeleteConsumerAddress(IOrganizationService service, Guid AddressGuid)
        {
            IoTAddressBookResultV1 objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            try
            {
                if (service != null)
                {
                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("statecode");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, AddressGuid);
                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResultV1 { AddressGuid = AddressGuid, StatusCode = 204, StatusDescription = "Address does not exist in D365." };
                        return objAddress;
                    }
                    else
                    {
                        if (entcoll.Entities[0].GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                        {
                            objAddress = new IoTAddressBookResultV1 { AddressGuid = AddressGuid, StatusCode = 204, StatusDescription = "Address is already deleted." };
                            return objAddress;
                        }
                    }
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <filter type='and'>
                                <condition attribute='hil_serviceaddress' operator='eq' value='{AddressGuid}' />
                            </filter>
                        </entity>
                        </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        objAddress = new IoTAddressBookResultV1 { AddressGuid = AddressGuid, StatusCode = 204, StatusDescription = $"Address is in use against JobId:{entcoll.Entities[0].GetAttributeValue<string>("msdyn_name")}.Can't be deleted." };
                        return objAddress;
                    }
                    else
                    {
                        Entity entity = new Entity("hil_address", AddressGuid);
                        entity["statecode"] = new OptionSetValue(1); //InActive
                        entity["statuscode"] = new OptionSetValue(2); //InActive
                        service.Update(entity);
                    }
                    objAddress = new IoTAddressBookResultV1 { AddressGuid = AddressGuid, StatusCode = 200, StatusDescription = "Success" };
                }
                else
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 503, StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResultV1 { StatusCode = 500, StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }
    }
}
