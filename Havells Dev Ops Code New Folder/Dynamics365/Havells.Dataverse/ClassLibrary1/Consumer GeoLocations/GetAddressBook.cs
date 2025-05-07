using Havells.Dataverse.CustomConnector.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_GeoLocations
{
    public class GetAddressBook :IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string JsonResponse = "";
            Guid CustomerGuid = Guid.Empty;
            string MobileNumber = null;
            bool isValidCustomerGuid = false;
            string Pincode = null;

            List<IoTAddressBookResult> lstAddressBookResult = new List<IoTAddressBookResult>();
            _tracingService.Trace("Start");

            if (context.InputParameters.Contains("CustomerGuid"))
            {
                _tracingService.Trace("CustomerGuid");
                isValidCustomerGuid = Guid.TryParse(Convert.ToString(context.InputParameters["CustomerGuid"]), out CustomerGuid);
            }
            if (context.InputParameters.Contains("MobileNumber"))
            {
                _tracingService.Trace("MobileNumber");

                MobileNumber = Convert.ToString(context.InputParameters["MobileNumber"]);
            }
            if (string.IsNullOrWhiteSpace(MobileNumber) && CustomerGuid == Guid.Empty)
            {
                lstAddressBookResult.Add(new IoTAddressBookResult
                {
                    StatusCode = "204",
                    StatusDescription = "Customer GUID/Mobile Number is required."
                });
                JsonResponse = JsonSerializer.Serialize(lstAddressBookResult);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            else if (string.IsNullOrWhiteSpace(MobileNumber))
            {
                if (!isValidCustomerGuid)
                {
                    lstAddressBookResult.Add(new IoTAddressBookResult
                    {
                        StatusCode = "204",
                        StatusDescription = "Invalid Customer GUID."
                    });
                    JsonResponse = JsonSerializer.Serialize(lstAddressBookResult);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            if (context.InputParameters.Contains("Pincode"))
            {
                Pincode = Convert.ToString(context.InputParameters["Pincode"]);
            }
            _tracingService.Trace("End");

            IoTAddressBook address = new IoTAddressBook { CustomerGuid = CustomerGuid, MobileNumber = MobileNumber, Pincode = Pincode };
            JsonResponse = JsonSerializer.Serialize(GetAddressBookDetails(service, address));
            _tracingService.Trace(JsonResponse);
            context.OutputParameters["data"] = JsonResponse;
        }
        public List<IoTAddressBookResult> GetAddressBookDetails(IOrganizationService service, IoTAddressBook address)
        {
            List<IoTAddressBookResult> addressList = new List<IoTAddressBookResult>();
            IoTAddressBookResult objAddress;
            QueryExpression query;
            EntityCollection collection;
            EntityReference erPincode = null;
            try
            {
                if (service != null)
                {
                    if (address.Pincode != null)
                    {
                        if (address.Pincode.Trim().Length == 6)
                        {
                            query = new QueryExpression("hil_pincode");
                            query.ColumnSet = new ColumnSet(false);
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, address.Pincode);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            EntityCollection entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count > 0)
                            {
                                erPincode = entcoll.Entities[0].ToEntityReference();
                            }
                            else
                            {
                                objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Invalid Pincode." };
                                addressList.Add(objAddress);
                                return addressList;
                            }
                        }
                        else
                        {
                            objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Invalid Pincode." };
                            addressList.Add(objAddress);
                            return addressList;
                        }
                    }

                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("hil_customer", "hil_addressid", "hil_street3", "hil_street2", "hil_street1", "hil_fulladdress",
                        "hil_pincode", "hil_area", "hil_businessgeo", "hil_addresstype", "hil_city", "hil_state");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    if (address.CustomerGuid != Guid.Empty)
                    {
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, address.CustomerGuid);
                    }
                    if (erPincode != null)
                    {
                        query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPincode.Id);
                    }
                    if (address.MobileNumber != null)
                    {
                        LinkEntity lnk = new LinkEntity("hil_address", "contact", "hil_customer", "contactid", JoinOperator.Inner);
                        lnk.LinkCriteria.AddCondition("mobilephone", ConditionOperator.Equal, address.MobileNumber);
                        query.LinkEntities.Add(lnk);
                    }
                    collection = service.RetrieveMultiple(query);

                    if (collection.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "No Address Book found." };
                        addressList.Add(objAddress);
                        return addressList;
                    }
                    if (collection.Entities != null && collection.Entities.Count > 0)
                    {
                        foreach (Entity item in collection.Entities)
                        {
                            objAddress = new IoTAddressBookResult();
                            if (item.Attributes.Contains("hil_street1"))
                            {
                                objAddress.AddressLine1 = item.GetAttributeValue<string>("hil_street1");
                            }
                            if (item.Attributes.Contains("hil_street2"))
                            {
                                objAddress.AddressLine2 = item.GetAttributeValue<string>("hil_street2");
                            }
                            if (item.Attributes.Contains("hil_street3"))
                            {
                                objAddress.AddressLine3 = item.GetAttributeValue<string>("hil_street3");
                            }
                            if (item.Attributes.Contains("hil_fulladdress"))
                            {
                                objAddress.FullAddress = item.GetAttributeValue<string>("hil_fulladdress");
                            }
                            if (item.Attributes.Contains("hil_addressid"))
                            {
                                objAddress.AddressGuid = item.Id;
                            }
                            if (item.Attributes.Contains("hil_businessgeo"))
                            {
                                objAddress.BizGeoGuid = item.GetAttributeValue<EntityReference>("hil_businessgeo").Id;
                            }
                            if (item.Attributes.Contains("hil_businessgeo"))
                            {
                                objAddress.BizGeoName = item.GetAttributeValue<EntityReference>("hil_businessgeo").Name;
                            }
                            if (item.Attributes.Contains("hil_pincode"))
                            {
                                objAddress.PINCodeGuid = item.GetAttributeValue<EntityReference>("hil_pincode").Id;
                            }
                            if (item.Attributes.Contains("hil_pincode"))
                            {
                                objAddress.PINCode = item.GetAttributeValue<EntityReference>("hil_pincode").Name;
                            }
                            if (item.Attributes.Contains("hil_area"))
                            {
                                objAddress.AreaGuid = item.GetAttributeValue<EntityReference>("hil_area").Id;
                            }
                            if (item.Attributes.Contains("hil_area"))
                            {
                                objAddress.Area = item.GetAttributeValue<EntityReference>("hil_area").Name;
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                OptionSetValue osv = item.GetAttributeValue<OptionSetValue>("hil_addresstype");
                                objAddress.AddressType = osv.Value == 1 ? "Permanent" : "Alternate";
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                objAddress.AddressTypeEnum = item.GetAttributeValue<OptionSetValue>("hil_addresstype").Value.ToString();
                            }
                            if (item.Attributes.Contains("hil_city"))
                            {
                                objAddress.CityName = item.GetAttributeValue<EntityReference>("hil_city").Name;
                            }
                            if (item.Attributes.Contains("hil_state"))
                            {
                                objAddress.StateName = item.GetAttributeValue<EntityReference>("hil_state").Name;
                            }
                            objAddress.CustomerGuid = address.CustomerGuid;
                            objAddress.MobileNumber = address.MobileNumber;
                            objAddress.StatusCode = "200";
                            objAddress.StatusDescription = "OK";
                            addressList.Add(objAddress);
                        }
                    }
                    return addressList;
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    addressList.Add(objAddress);
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                addressList.Add(objAddress);
            }
            return addressList;
        }

    }
}
