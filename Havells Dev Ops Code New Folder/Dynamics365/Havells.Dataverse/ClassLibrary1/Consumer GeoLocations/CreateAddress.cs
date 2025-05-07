using Havells.Dataverse.CustomConnector.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class CreateAddress : IPlugin
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
            Guid PINCodeGuid = Guid.Empty;
            string AddressLine1 = string.Empty;

            string AddressTypeEnum = string.Empty;
            Guid AreaGuid = Guid.Empty;
            string AddressLine2 = string.Empty;
            string AddressLine3 = string.Empty;

            StringBuilder errorMessage = new StringBuilder();
            bool IsValidRequest = true;
            string JsonResponse = "";
            if (context.InputParameters.Contains("CustomerGuid") && context.InputParameters["CustomerGuid"] is string
                && context.InputParameters.Contains("PINCodeGuid") && context.InputParameters["PINCodeGuid"] is string
                && context.InputParameters.Contains("AddressLine1") && context.InputParameters["AddressLine1"] is string)
            {
                bool isValidGuid = Guid.TryParse(Convert.ToString(context.InputParameters["CustomerGuid"]), out CustomerGuid);
                if (!isValidGuid)
                {
                    errorMessage.AppendLine("Invalid Customer GuId.");
                    IsValidRequest = false;
                }
                isValidGuid = Guid.TryParse(Convert.ToString(context.InputParameters["PINCodeGuid"]), out PINCodeGuid);
                if (!isValidGuid)
                {
                    errorMessage.AppendLine("Invalid PINCode GuId.");
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
                    PINCodeGuid = PINCodeGuid,
                    AddressLine1 = AddressLine1,
                    AddressLine2 = AddressLine2,
                    AddressLine3 = AddressLine3,
                    AddressTypeEnum = AddressTypeEnum,
                    AreaGuid = AreaGuid
                };
                JsonResponse = JsonSerializer.Serialize(CreateConsumerAddress(service, address));
                _tracingService.Trace(JsonResponse);
                context.OutputParameters["data"] = JsonResponse;
            }
        }
        public IoTAddressBookResult CreateConsumerAddress(IOrganizationService service, IoTAddressBookResult addressData)
        {
            IoTAddressBookResult objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            EntityReference businessGeo = null;

            try
            {
                if (service != null)
                {
                    int addressType;
                    if (string.IsNullOrWhiteSpace(addressData.AddressTypeEnum))
                    {
                        query = new QueryExpression("hil_address");
                        query.ColumnSet = new ColumnSet("hil_addressid");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, addressData.CustomerGuid);
                        entcoll = service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count > 0) { addressType = 2; }
                        else { addressType = 1; }
                    }
                    else
                    {
                        if (addressData.AddressTypeEnum == "1")
                        {
                            addressType = 1;
                        }
                        else
                        {
                            addressType = 2;
                        }
                    }

                    query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_pincode");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, addressData.PINCodeGuid);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active
                    if (addressData.AreaGuid != Guid.Empty)
                    {
                        query.Criteria.AddCondition("hil_area", ConditionOperator.Equal, addressData.AreaGuid);
                    }
                    query.AddOrder("createdon", OrderType.Ascending);
                    query.TopCount = 1;

                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count > 0)
                    {
                        businessGeo = entcoll.Entities[0].ToEntityReference();
                    }
                    else
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "PIN Code or Area does not exist." };
                        return objAddress;
                    }
                    Entity entObj = new Entity("hil_address");
                    entObj["hil_street1"] = addressData.AddressLine1;
                    if (!string.IsNullOrWhiteSpace(addressData.AddressLine2))
                    {
                        entObj["hil_street2"] = addressData.AddressLine2;
                    }
                    if (!string.IsNullOrWhiteSpace(addressData.AddressLine3))
                    {
                        entObj["hil_street3"] = addressData.AddressLine3;
                    }
                    entObj["hil_customer"] = new EntityReference("contact", addressData.CustomerGuid);
                    if (businessGeo != null)
                    {
                        entObj["hil_businessgeo"] = businessGeo;
                    }
                    entObj["hil_addresstype"] = new OptionSetValue(addressType);
                    objAddress = new IoTAddressBookResult();
                    objAddress.AddressGuid = service.Create(entObj);
                    if (objAddress.AddressGuid != Guid.Empty)
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.StatusCode = "200";
                        objAddress.StatusDescription = "OK.";
                    }
                    else
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.StatusCode = "204";
                        objAddress.StatusDescription = "Something went wrong.";
                    }
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
