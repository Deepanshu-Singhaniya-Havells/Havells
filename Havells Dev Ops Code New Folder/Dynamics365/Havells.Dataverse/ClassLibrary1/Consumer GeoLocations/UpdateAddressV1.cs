using Havells.Dataverse.CustomConnector.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class UpdateAddressV1 : IPlugin
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

            int AddressTypeEnum = 3;
            Guid AreaGuid = Guid.Empty;
            string AddressLine2 = string.Empty;
            string AddressLine3 = string.Empty;
            bool IsDefault = false;

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
                    JsonResponse = JsonSerializer.Serialize(new IoTAddressBookResultV1
                    {
                        StatusCode = 204,
                        StatusDescription = errorMessage.ToString(),
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (context.InputParameters.Contains("AddressTypeEnum"))
                {
                    int.TryParse(Convert.ToString(context.InputParameters["AddressTypeEnum"]), out AddressTypeEnum);
                }
                Guid.TryParse(Convert.ToString(context.InputParameters["AreaGuid"]), out AreaGuid);
                IsDefault = Convert.ToBoolean(context.InputParameters["IsDefault"]);
                AddressLine2 = Convert.ToString(context.InputParameters["AddressLine2"]);
                AddressLine3 = Convert.ToString(context.InputParameters["AddressLine3"]);

                IoTAddressBookResultV1 address = new IoTAddressBookResultV1
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

        public IoTAddressBookResultV1 UpdateConsumerAddress(IOrganizationService service, IoTAddressBookResultV1 addressData)
        {
            IoTAddressBookResultV1 objAddress = null;
            try
            {
                if (addressData.AddressGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address GUID is required." };
                    return objAddress;
                }
                if (addressData.CustomerGuid == Guid.Empty)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Customer GUID is required." };
                    return objAddress;
                }
                if (addressData.AddressLine1.Trim().Length == 0)
                {
                    objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "Address Line 1 is required." };
                    return objAddress;
                }

                if (service != null)
                {
                    Entity address = service.Retrieve("hil_address", addressData.AddressGuid, new ColumnSet("hil_customer"));

                    if (address.Contains("hil_customer"))
                    {
                        if (address.GetAttributeValue<EntityReference>("hil_customer").Id != addressData.CustomerGuid)
                        {
                            objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = "Address ID does not belong to Customer." };
                            return objAddress;
                        }
                    }
                    else
                    {
                        objAddress = new IoTAddressBookResultV1 { AddressGuid = addressData.AddressGuid, StatusCode = 204, StatusDescription = "Address ID does not belong to Customer." };
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
                    int[] AddressType = { 1, 2, 3 };
                    int AddressTypeEnum = AddressType.Contains(addressData.AddressTypeEnum) ? addressData.AddressTypeEnum : 3;
                    entObj["hil_addresstype"] = new OptionSetValue(AddressTypeEnum);
                    if (addressData.IsDefault)
                    {
                        RemoveDefaultAddress(service, addressData.CustomerGuid);
                    }
                    entObj["hil_isdefault"] = addressData.IsDefault;
                    service.Update(entObj);

                    objAddress = new IoTAddressBookResultV1();
                    objAddress.AddressGuid = addressData.AddressGuid;
                    objAddress.CustomerGuid = addressData.CustomerGuid;
                    objAddress.AddressLine1 = addressData.AddressLine1;
                    objAddress.AddressLine2 = addressData.AddressLine2;
                    objAddress.AddressLine3 = addressData.AddressLine3;
                    objAddress.AddressTypeEnum = AddressTypeEnum;
                    if (AddressTypeEnum == 1)
                    {
                        objAddress.AddressType = "Home";
                    }
                    else if (AddressTypeEnum == 2)
                    {
                        objAddress.AddressType = "Office";
                    }
                    else
                    {
                        objAddress.AddressType = "Other";
                    }
                    objAddress.IsDefault = addressData.IsDefault;
                    objAddress.StatusCode = 200;
                    objAddress.StatusDescription = "OK.";
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
        public void RemoveDefaultAddress(IOrganizationService service, Guid CustomerGuid)
        {
            QueryExpression query = new QueryExpression("hil_address");
            query.ColumnSet = new ColumnSet("hil_customer", "hil_addressid", "hil_addresstype", "hil_isdefault");
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("hil_isdefault", ConditionOperator.Equal, true);
            query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerGuid);
            EntityCollection collection = service.RetrieveMultiple(query);

            if (collection.Entities.Count > 0)
            {
                foreach (var item in collection.Entities)
                {
                    Entity entity = new Entity("hil_address", item.Id);
                    entity["hil_isdefault"] = false;
                    service.Update(entity);
                }
            }
        }
    }
}
