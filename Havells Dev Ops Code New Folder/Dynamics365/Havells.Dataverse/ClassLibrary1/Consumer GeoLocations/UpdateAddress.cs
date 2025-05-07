using Havells.Dataverse.CustomConnector.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class UpdateAddress : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            Guid CustomerGuid = Guid.Empty;
            Guid AddressGuid = Guid.Empty;
            string AddressLine1 = string.Empty;

            string AddressTypeEnum = string.Empty;
            Guid AreaGuid = Guid.Empty;
            string AddressLine2 = string.Empty;
            string AddressLine3 = string.Empty;

            StringBuilder errorMessage = new StringBuilder();
            bool IsValidRequest = true;
            string JsonResponse = "";
            if (context.InputParameters.Contains("CustomerGuid") && context.InputParameters["CustomerGuid"] is string
                && context.InputParameters.Contains("AddressGuid") && context.InputParameters["AddressGuid"] is string
                && context.InputParameters.Contains("AddressLine1") && context.InputParameters["AddressLine1"] is string)
            {
                bool isValidGuid = Guid.TryParse(Convert.ToString(context.InputParameters["CustomerGuid"]), out CustomerGuid);
                if (!isValidGuid)
                {
                    errorMessage.AppendLine("Invalid Customer Guid.");
                    IsValidRequest = false;
                }
                isValidGuid = Guid.TryParse(Convert.ToString(context.InputParameters["AddressGuid"]), out AddressGuid);
                if (!isValidGuid)
                {
                    errorMessage.AppendLine("Invalid Address Guid.");
                    IsValidRequest = false;
                }
                AddressLine1 = Convert.ToString(context.InputParameters["AddressLine1"]);
                if (string.IsNullOrWhiteSpace(AddressLine1))
                {
                    errorMessage.AppendLine("Address Line 1 is required.");
                    IsValidRequest = false;
                }
                if (!IsValidRequest)
                {
                    JsonResponse = JsonSerializer.Serialize(new IoTAddressBookResult
                    {
                        StatusCode = "204",
                        StatusDescription = errorMessage.ToString(),
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (context.InputParameters.Contains("AddressTypeEnum"))
                {
                    AddressTypeEnum = Convert.ToString(context.InputParameters["AddressTypeEnum"]);
                }
                isValidGuid = Guid.TryParse(Convert.ToString(context.InputParameters["AreaGuid"]), out AreaGuid);
                AddressLine2 = Convert.ToString(context.InputParameters["AddressLine2"]);
                AddressLine3 = Convert.ToString(context.InputParameters["AddressLine3"]);

                IoTAddressBookResult address = new IoTAddressBookResult
                {
                    CustomerGuid = CustomerGuid,
                    AddressGuid = AddressGuid,
                    AddressLine1 = AddressLine1,
                    AddressLine2 = AddressLine2,
                    AddressLine3 = AddressLine3,
                    AddressTypeEnum = AddressTypeEnum,
                    AreaGuid = AreaGuid
                };
                JsonResponse = JsonSerializer.Serialize(UpdateConsumerAddress(service, address));
                _tracingService.Trace(JsonResponse);
                context.OutputParameters["data"] = JsonResponse;
            }
        }
        public IoTAddressBookResult UpdateConsumerAddress(IOrganizationService service, IoTAddressBookResult addressData)
        {
            IoTAddressBookResult objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            try
            {
                if (service != null)
                {
                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("hil_addressid");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, addressData.CustomerGuid);
                    query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, addressData.AddressGuid);
                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Address ID does not belong to Customer." };
                        return objAddress;
                    }
                    Entity entObj = new Entity("hil_address", addressData.AddressGuid);
                    entObj["hil_street1"] = addressData.AddressLine1;
                    if (addressData.AddressLine2 != null)
                    {
                        entObj["hil_street2"] = addressData.AddressLine2;
                    }
                    if (addressData.AddressLine3 != null)
                    {
                        entObj["hil_street3"] = addressData.AddressLine3;
                    }
                    entObj["hil_customer"] = new EntityReference("contact", addressData.CustomerGuid);
                    if (addressData.AddressTypeEnum == "1")
                    {
                        entObj["hil_addresstype"] = new OptionSetValue(1); //Permanent
                    }
                    else
                    {
                        entObj["hil_addresstype"] = new OptionSetValue(2); //Alternate
                    }
                    service.Update(entObj);
                    objAddress = new IoTAddressBookResult();
                    objAddress.AddressGuid = addressData.AddressGuid;
                    objAddress.CustomerGuid = addressData.CustomerGuid;
                    objAddress.StatusCode = "200";
                    objAddress.StatusDescription = "OK.";
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objAddress;
        }
    }
}
