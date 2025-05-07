using Havells.Dataverse.CustomConnector.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class CreateAddressV1 : IPlugin
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
            string[] AddressType = { "1", "2", "3", "4", "5" };
            int AddressTypeEnum = 3;
            Guid AreaGuid = Guid.Empty;
            string AddressLine2 = string.Empty;
            string AddressLine3 = string.Empty;
            bool IsDefaultAddress = false;

            StringBuilder errorMessage = new StringBuilder();
            bool IsValidRequest = true;
            string JsonResponse = "";
            AddressLine1 = Convert.ToString(context.InputParameters["AddressLine1"]);
            string Customerid = Convert.ToString(context.InputParameters["CustomerGuid"]);
            string PINCodeid = Convert.ToString(context.InputParameters["PINCodeGuid"]);
            string AddressTypeValue = Convert.ToString(context.InputParameters["AddressTypeEnum"]);
            if (!APValidate.IsvalidGuid(Customerid) || Customerid == Guid.Empty.ToString())
            {
                string msg = string.IsNullOrWhiteSpace(Customerid) || Customerid == Guid.Empty.ToString() ? "CustomerGuid required." : "Invalid CustomerGuid.";
                errorMessage.AppendLine(msg);
                IsValidRequest = false;
            }
            if (!APValidate.IsvalidGuid(PINCodeid) || PINCodeid == Guid.Empty.ToString())
            {
                string msg = string.IsNullOrWhiteSpace(PINCodeid) || PINCodeid == Guid.Empty.ToString() ? "PINCodeGuId required." : "Invalid PINCodeGuId.";
                errorMessage.AppendLine(msg);
                IsValidRequest = false;
            }
            if (string.IsNullOrWhiteSpace(AddressTypeValue) || !AddressType.Contains(AddressTypeValue))
            {
                string msg = string.IsNullOrWhiteSpace(AddressTypeValue) ? "AddressTypeEnum required." : "Invalid AddressTypeEnum.";
                errorMessage.AppendLine(msg);
                IsValidRequest = false;
            }
            if (string.IsNullOrWhiteSpace(AddressLine1.Trim()))
            {
                errorMessage.AppendLine("AddressLine1 is required.");
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

            Guid.TryParse(Convert.ToString(context.InputParameters["AreaGuid"]), out AreaGuid);
            IsDefaultAddress = Convert.ToBoolean(context.InputParameters["IsDefault"]);
            AddressLine2 = Convert.ToString(context.InputParameters["AddressLine2"]);
            AddressLine3 = Convert.ToString(context.InputParameters["AddressLine3"]);

            IoTAddressBookResultV1 address = new IoTAddressBookResultV1
            {
                CustomerGuid = new Guid(Customerid),
                PINCodeGuid = new Guid(PINCodeid),
                AddressLine1 = AddressLine1,
                AddressLine2 = AddressLine2,
                AddressLine3 = AddressLine3,
                AddressTypeEnum = Convert.ToInt32(AddressTypeValue),
                AreaGuid = AreaGuid,
                IsDefault = IsDefaultAddress
            };
            JsonResponse = JsonSerializer.Serialize(CreateConsumerAddress(service, address));
            _tracingService.Trace(JsonResponse);
            context.OutputParameters["data"] = JsonResponse;
        }

        public IoTAddressBookResultV1 CreateConsumerAddress(IOrganizationService service, IoTAddressBookResultV1 addressData)
        {
            IoTAddressBookResultV1 objAddress = null;
            QueryExpression query;
            EntityCollection entcoll;
            EntityReference businessGeo = null;
            EntityReference district = null;

            try
            {
                if (service != null)
                {
                    query = new QueryExpression("hil_businessmapping");
                    query.ColumnSet = new ColumnSet("hil_pincode", "hil_district");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, addressData.PINCodeGuid);
                    query.Criteria.AddCondition("hil_district", ConditionOperator.NotNull);
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
                        district = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_district");
                    }
                    else
                    {
                        objAddress = new IoTAddressBookResultV1 { StatusCode = 204, StatusDescription = "PIN Code or Area does not exist." };
                        return objAddress;
                    }
                    Entity entObj = new Entity("hil_address");
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
                    if (businessGeo != null)
                    {
                        entObj["hil_businessgeo"] = businessGeo;
                    }
                    if (district != null)
                    {
                        entObj["hil_district"] = district;
                    }
                    int[] intAddressType = { 1, 2, 3 };
                    int AddressTypeEnum = intAddressType.Contains(addressData.AddressTypeEnum) ? addressData.AddressTypeEnum : 3;
                    entObj["hil_addresstype"] = new OptionSetValue(AddressTypeEnum);
                    if (addressData.IsDefault)
                    {
                        RemoveDefaultAddress(service, addressData.CustomerGuid);
                        entObj["hil_isdefault"] = addressData.IsDefault;
                    }
                    objAddress = new IoTAddressBookResultV1();
                    objAddress.AddressGuid = service.Create(entObj);
                    if (objAddress.AddressGuid != Guid.Empty)
                    {
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.AddressLine1 = addressData.AddressLine1;
                        objAddress.AddressLine2 = addressData.AddressLine2;
                        objAddress.AddressLine3 = addressData.AddressLine3;
                        objAddress.PINCodeGuid = addressData.PINCodeGuid;
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
                        objAddress.CustomerGuid = addressData.CustomerGuid;
                        objAddress.StatusCode = 204;
                        objAddress.StatusDescription = "Something went wrong.";
                    }
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
