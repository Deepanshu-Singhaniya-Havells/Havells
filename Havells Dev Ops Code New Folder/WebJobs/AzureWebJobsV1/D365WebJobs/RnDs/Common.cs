using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Text;
using Newtonsoft.Json;
using System.Net;
using Microsoft.Xrm.Sdk.Query;

namespace D365WebJobs
{
    public class Common
    {
        public static void SetFullAddress(Entity entity, String sMessage, IOrganizationService service)
        {
            try
            {
                ////tracingService.Trace("Common1");

                #region Create
                if (sMessage == "CREATE")
                {

                    hil_address MyAddress = new hil_address();
                    Contact enContact = entity.ToEntity<Contact>();
                    String sFullAddress = String.Empty;

                    Contact upContact = new Contact();
                    upContact.ContactId = entity.Id;

                    if (enContact.Address1_Line1 != null)
                    {
                        MyAddress.hil_Street1 = enContact.Address1_Line1;
                        ////tracingService.Trace("Common2");
                        sFullAddress = enContact.Address1_Line1 + " ";
                    }
                    if (enContact.Address1_Line2 != null)
                    {
                        MyAddress.hil_Street2 = enContact.Address1_Line2;
                        ////tracingService.Trace("Common3");
                        sFullAddress += enContact.Address1_Line2 + " ";
                    }
                    if (enContact.Address1_Line3 != null)
                    {
                        MyAddress.hil_Street3 = enContact.Address1_Line3;
                        ////tracingService.Trace("Common4");
                        sFullAddress += enContact.Address1_Line3 + " ";
                    }
                    if (enContact.hil_city != null)
                    {
                        MyAddress.hil_CIty = enContact.hil_city;
                        upContact.hil_TextCity = enContact.hil_city.Name;
                        ////tracingService.Trace("Common5");
                        sFullAddress += enContact.hil_city.Name + " ";
                    }
                    if (enContact.hil_district != null)
                    {
                        MyAddress.hil_District = enContact.hil_district;
                        upContact.hil_TextDistrict = enContact.hil_district.Name;
                        ////tracingService.Trace("Common6");
                        sFullAddress += enContact.hil_district.Name + " ";
                    }
                    if (enContact.hil_pincode != null)
                    {
                        MyAddress.hil_PinCode = enContact.hil_pincode;
                        upContact.hil_TextPinCode = enContact.hil_pincode.Name;
                        ////tracingService.Trace("Common7");
                        sFullAddress += enContact.hil_pincode.Name + " ";
                    }
                    else return;
                    if (enContact.hil_area != null)
                    {

                        upContact.hil_TextArea = enContact.hil_area.Name;
                        ////tracingService.Trace("Common8");
                        sFullAddress += enContact.hil_area.Name + " ";
                    }
                    if (enContact.hil_state != null)
                    {
                        MyAddress.hil_State = enContact.hil_state;
                        upContact.hil_TextState = enContact.hil_state.Name;
                        ////tracingService.Trace("Common9");
                        sFullAddress += enContact.hil_state.Name + " ";
                    }
                    if (enContact.hil_geomappingpincode != null)
                    {
                        MyAddress.hil_IPGeo = enContact.hil_geomappingpincode;
                    }

                    ////tracingService.Trace("Common10" + sFullAddress);
                    upContact.hil_FullAddress = sFullAddress;
                    service.Update(upContact);

                    MyAddress.hil_FullAddress = sFullAddress;
                    MyAddress.hil_AddressType = new OptionSetValue(1);
                    MyAddress.hil_name = "PermanentAddress";
                    MyAddress.hil_Customer = new EntityReference(Contact.EntityLogicalName, enContact.Id);
                    service.Create(MyAddress);

                    #region CreateShippingAddress
                    hil_address ShippingAddress = new hil_address();
                    String sShipFullAddress = String.Empty;
                    if (enContact.Address2_Line1 != String.Empty)
                    {
                        ShippingAddress.hil_Street1 = enContact.Address2_Line1;
                        sShipFullAddress = enContact.Address2_Line1 + " ";
                    }
                    if (enContact.Address2_Line2 != String.Empty)
                    {
                        ShippingAddress.hil_Street2 = enContact.Address2_Line2;
                        sShipFullAddress += enContact.Address2_Line2 + " ";
                    }
                    if (enContact.Address2_Line3 != String.Empty)
                    {
                        ShippingAddress.hil_Street3 = enContact.Address2_Line3;
                        sShipFullAddress += enContact.Address2_Line3 + " ";
                    }
                    if (enContact.hil_ShippingCity != null)
                    {
                        ShippingAddress.hil_CIty = enContact.hil_ShippingCity;
                        sShipFullAddress += enContact.hil_ShippingCity.Name + " ";
                    }
                    if (enContact.hil_ShippingDistrict != null)
                    {
                        ShippingAddress.hil_District = enContact.hil_ShippingDistrict;
                        sShipFullAddress += enContact.hil_ShippingDistrict.Name + " ";
                    }
                    if (enContact.hil_ShippingPinCode != null)
                    {
                        ShippingAddress.hil_PinCode = enContact.hil_ShippingPinCode;
                        sShipFullAddress += enContact.hil_ShippingPinCode.Name + " ";
                    }
                    else return;
                    if (enContact.hil_ShippingState != null)
                    {
                        ShippingAddress.hil_State = enContact.hil_ShippingState;
                        sShipFullAddress += enContact.hil_ShippingState.Name + " ";
                    }
                    if (enContact.hil_GeoLocatorShipping != null)
                    {
                        ShippingAddress.hil_IPGeo = enContact.hil_GeoLocatorShipping;
                    }
                    ShippingAddress.hil_FullAddress = sShipFullAddress;

                    ShippingAddress.hil_AddressType = new OptionSetValue(2);
                    ShippingAddress.hil_name = "CurrentAddress";
                    ShippingAddress.hil_Customer = new EntityReference(Contact.EntityLogicalName, enContact.Id);
                    service.Create(ShippingAddress);

                    #endregion
                }
                #endregion
                #region Update
                else if (sMessage == "UPDATE")
                {
                    Guid fsPrimaryAddressId = Guid.Empty;
                    Guid fsShippingAddressId = Guid.Empty;

                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        var obj = from _Address in orgContext.CreateQuery<hil_address>()
                                  where (_Address.hil_AddressType.Value == 1 || _Address.hil_AddressType.Value == 2)
                                  && _Address.hil_Customer.Id == entity.Id
                                  select new
                                  {
                                      _Address.hil_addressId
                                  ,
                                      _Address.hil_AddressType
                                  };
                        Int32 count = Enumerable.Count(obj);
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_AddressType.Value == 1)
                                fsPrimaryAddressId = iobj.hil_addressId.Value;
                            if (iobj.hil_AddressType.Value == 2)
                                fsShippingAddressId = iobj.hil_addressId.Value;
                        }
                    }


                    hil_address MyAddress = new hil_address();

                    Contact enContact = entity.ToEntity<Contact>();
                    String sFullAddress = String.Empty;

                    Contact upContact = new Contact();
                    upContact.ContactId = entity.Id;

                    if (enContact.Address1_Line1 != null)
                    {
                        MyAddress.hil_Street1 = enContact.Address1_Line1;
                        ////tracingService.Trace("Common2");
                        sFullAddress = enContact.Address1_Line1 + " ";
                    }
                    if (enContact.Address1_Line2 != null)
                    {
                        MyAddress.hil_Street2 = enContact.Address1_Line2;
                        ////tracingService.Trace("Common3");
                        sFullAddress += enContact.Address1_Line2 + " ";
                    }
                    if (enContact.Address1_Line3 != null)
                    {
                        MyAddress.hil_Street3 = enContact.Address1_Line3;
                        ////tracingService.Trace("Common4");
                        sFullAddress += enContact.Address1_Line3 + " ";
                    }
                    if (enContact.hil_city != null)
                    {
                        MyAddress.hil_CIty = enContact.hil_city;
                        upContact.hil_TextCity = enContact.hil_city.Name;
                        ////tracingService.Trace("Common5");
                        sFullAddress += enContact.hil_city.Name + " ";
                    }
                    if (enContact.hil_district != null)
                    {
                        MyAddress.hil_District = enContact.hil_district;
                        upContact.hil_TextDistrict = enContact.hil_district.Name;
                        ////tracingService.Trace("Common6");
                        sFullAddress += enContact.hil_district.Name + " ";
                    }
                    if (enContact.hil_pincode != null)
                    {
                        MyAddress.hil_PinCode = enContact.hil_pincode;
                        upContact.hil_TextPinCode = enContact.hil_pincode.Name;
                        ////tracingService.Trace("Common7");
                        sFullAddress += enContact.hil_pincode.Name + " ";
                    }
                    else return;
                    if (enContact.hil_area != null)
                    {

                        upContact.hil_TextArea = enContact.hil_area.Name;
                        ////tracingService.Trace("Common8");
                        sFullAddress += enContact.hil_area.Name + " ";
                    }
                    if (enContact.hil_state != null)
                    {
                        MyAddress.hil_State = enContact.hil_state;
                        upContact.hil_TextState = enContact.hil_state.Name;
                        ////tracingService.Trace("Common9");
                        sFullAddress += enContact.hil_state.Name + " ";
                    }
                    if (enContact.hil_geomappingpincode != null)
                    {
                        MyAddress.hil_IPGeo = enContact.hil_geomappingpincode;
                    }

                    ////tracingService.Trace("Common10" + sFullAddress);
                    upContact.hil_FullAddress = sFullAddress;
                    service.Update(upContact);

                    MyAddress.hil_FullAddress = sFullAddress;
                    MyAddress.hil_AddressType = new OptionSetValue(1);
                    if (fsPrimaryAddressId != Guid.Empty)
                    {
                        MyAddress.hil_addressId = fsPrimaryAddressId;
                        service.Update(MyAddress);
                    }
                    else
                    {
                        MyAddress.hil_name = "PermanentAddress";
                        MyAddress.hil_Customer = new EntityReference(Contact.EntityLogicalName, enContact.Id);
                        service.Create(MyAddress);
                    }
                    #region CreateShippingAddress
                    hil_address ShippingAddress = new hil_address();
                    String sShipFullAddress = String.Empty;
                    if (enContact.Address2_Line1 != String.Empty)
                    {
                        ShippingAddress.hil_Street1 = enContact.Address2_Line1;
                        sShipFullAddress = enContact.Address2_Line1 + " ";
                    }
                    if (enContact.Address2_Line2 != String.Empty)
                    {
                        ShippingAddress.hil_Street2 = enContact.Address2_Line2;
                        sShipFullAddress += enContact.Address2_Line2 + " ";
                    }
                    if (enContact.Address2_Line3 != String.Empty)
                    {
                        ShippingAddress.hil_Street3 = enContact.Address2_Line3;
                        sShipFullAddress += enContact.Address2_Line3 + " ";
                    }
                    if (enContact.hil_ShippingCity != null)
                    {
                        ShippingAddress.hil_CIty = enContact.hil_ShippingCity;
                        sShipFullAddress += enContact.hil_ShippingCity.Name + " ";
                    }
                    if (enContact.hil_ShippingDistrict != null)
                    {
                        ShippingAddress.hil_District = enContact.hil_ShippingDistrict;
                        sShipFullAddress += enContact.hil_ShippingDistrict.Name + " ";
                    }
                    if (enContact.hil_ShippingPinCode != null)
                    {
                        ShippingAddress.hil_PinCode = enContact.hil_ShippingPinCode;
                        sShipFullAddress += enContact.hil_ShippingPinCode.Name + " ";
                    }
                    else return;
                    if (enContact.hil_ShippingState != null)
                    {
                        ShippingAddress.hil_State = enContact.hil_ShippingState;
                        sShipFullAddress += enContact.hil_ShippingState.Name + " ";
                    }
                    if (enContact.hil_GeoLocatorShipping != null)
                    {
                        ShippingAddress.hil_IPGeo = enContact.hil_GeoLocatorShipping;
                    }


                    ShippingAddress.hil_FullAddress = sShipFullAddress;
                    ShippingAddress.hil_AddressType = new OptionSetValue(2);
                    if (fsShippingAddressId != Guid.Empty)
                    {
                        ShippingAddress.hil_addressId = fsShippingAddressId;
                        service.Update(ShippingAddress);
                    }
                    
                   else 
                    {
                        ShippingAddress.hil_name = "CurrentAddress";
                        ShippingAddress.hil_Customer = new EntityReference(Contact.EntityLogicalName, enContact.Id);
                       
                        service.Create(ShippingAddress);
                    }
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ContactEn.GetFullAddress" + ex.Message);
            }
        }

        public static string getFormatedDate(string date, ITracingService tracingService)
        {
            //tracingService.Trace("Date " + date);

            try
            {
                //tracingService.Trace("Date " + date);
                string[] dateTime = date.Split(' ');
                string[] dateOnly = dateTime[0].Split('/');
                string formatedDate = dateOnly[2] + "-" + dateOnly[0] + "-" + dateOnly[1] + " " + dateTime[1];
                return formatedDate;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error12 " + ex.Message);
            }
        }
        public static void CallAPIInsertUpdateCustomer(IOrganizationService service, Entity contactEntity, ITracingService tracingService)
        {
            try
            {
                var payload = getContactDetails(service, contactEntity, tracingService);

                //tracingService.Trace("CallAPIInsertUpdateCustomer Line 2");
                String requestStr = "[" + JsonConvert.SerializeObject(payload) + "]";
                Integration integration = IntegrationConfiguration(service, "ConsumerToMDMSync", tracingService);
                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri;
                String response = string.Empty;
                //tracingService.Trace("***************  requestStr " + requestStr);
                using (WebClient webClient = new WebClient())
                {
                    webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + _authInfo;
                    webClient.Headers[HttpRequestHeader.ContentType] = "application/json";
                    response = webClient.UploadString(sUrl, requestStr);
                }
                //tracingService.Trace(response);
                response = response.Replace('[', ' ').Replace(']', ' ');
                var responseContact = JsonConvert.DeserializeObject<InsertUpdateCustomerResponse>(response);
                //tracingService.Trace("Updating Entity");
                Entity _contact = new Entity("contact");
                _contact["hil_mdmid"] = responseContact.MDMId != "" ? responseContact.MDMId : null;
                _contact["hil_errormsg"] = responseContact.ErrorMSG != "" ? responseContact.ErrorMSG : null;
                _contact.Id = contactEntity.Id;
                service.Update(_contact);
                //tracingService.Trace("Done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            //Guid contactId = new Guid("9eef1e76-b636-eb11-a813-0022486e64c4");

        }
        public static InsertUpdateCustomerPayLoad getContactDetails(IOrganizationService service, Entity contactEntity, ITracingService tracingService)
        {
            //tracingService.Trace("getContactDetails");
            InsertUpdateCustomerPayLoad payLoad = new InsertUpdateCustomerPayLoad();
            try
            {
                //contactEntity = service.Retrieve("contact", contactEntity.Id, new ColumnSet(true));

                payLoad.Salutation = contactEntity.Contains("hil_salutation") ? contactEntity.GetAttributeValue<OptionSetValue>("hil_salutation").Value.ToString() : null;
                //tracingService.Trace("hil_salutation");
                payLoad.FirstName = contactEntity.Contains("firstname") ? contactEntity.GetAttributeValue<string>("firstname") : null;
                payLoad.MiddleName = contactEntity.Contains("middlename") ? contactEntity.GetAttributeValue<string>("middlename") : null;
                payLoad.LastName = contactEntity.Contains("middlename") ? contactEntity.GetAttributeValue<string>("lastname") : null;
                payLoad.FullName = contactEntity.Contains("fullname") ? contactEntity.GetAttributeValue<string>("fullname") : null;
                payLoad.EmailAddress = contactEntity.Contains("emailaddress1") ? contactEntity.GetAttributeValue<string>("emailaddress1") : null;
                payLoad.Consent = contactEntity.Contains("hil_consent") ? contactEntity.GetAttributeValue<bool>("hil_consent") ? "1" : "0" : null;
                //tracingService.Trace("hil_consent");
                payLoad.ConsentName = contactEntity.Contains("hil_consent") ? contactEntity.GetAttributeValue<bool>("hil_consent") ? "Yes" : "No" : null;
                payLoad.SourceOfCreation = contactEntity.Contains("hil_consumersource") ? contactEntity.GetAttributeValue<OptionSetValue>("hil_consumersource").Value.ToString() : null;
                //tracingService.Trace("hil_consumersource");
                payLoad.SourceOfCreationName = contactEntity.Contains("hil_consumersource") ? contactEntity.FormattedValues["hil_consumersource"] : null;
                payLoad.MobileNumber = contactEntity.Contains("mobilephone") ? contactEntity.GetAttributeValue<string>("mobilephone") : null;
                payLoad.Gender = contactEntity.Contains("hil_gender") ? contactEntity.GetAttributeValue<bool>("hil_gender") ? "1" : "0" : null;
                payLoad.GenderName = contactEntity.Contains("hil_gender") ? contactEntity.GetAttributeValue<bool>("hil_gender") ? "female" : "male" : null;
                //tracingService.Trace("hil_gender");

                payLoad.DateOfBirth = contactEntity.Contains("hil_dateofbirth") ? getFormatedDate(contactEntity.GetAttributeValue<DateTime>("hil_dateofbirth").AddMinutes(330).ToString(), tracingService) : null;
                //tracingService.Trace("hil_dateofbirth");
                payLoad.DateOfAnniversary = contactEntity.Contains("hil_dateofanniversary") ? getFormatedDate(contactEntity.GetAttributeValue<DateTime>("hil_dateofanniversary").AddMinutes(330).ToString(), tracingService) : null;
                //tracingService.Trace("hil_dateofanniversary");

                //tracingService.Trace("Created on " + contactEntity.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                //tracingService.Trace("Modified on " + contactEntity.GetAttributeValue<DateTime>("modifiedon").AddMinutes(330).ToString());
                payLoad.AlternateNumber = contactEntity.Contains("address1_telephone3") ? contactEntity.GetAttributeValue<string>("address1_telephone3") : null;

                //tracingService.Trace("8");
                payLoad.ContactId = contactEntity.Id.ToString();
                //tracingService.Trace("1");
                payLoad.Createdby = contactEntity.GetAttributeValue<EntityReference>("createdby").Id.ToString();
                //tracingService.Trace("2");
                payLoad.CreatedbyName = contactEntity.GetAttributeValue<EntityReference>("createdby").Name.ToString();
                //tracingService.Trace("3");
                payLoad.CreatedOn = getFormatedDate(contactEntity.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString(), tracingService);
                //tracingService.Trace("4");
                payLoad.Modifiedby = contactEntity.GetAttributeValue<EntityReference>("modifiedby").Id.ToString();
                //tracingService.Trace("5");
                payLoad.ModifiedByName = contactEntity.GetAttributeValue<EntityReference>("modifiedby").Name.ToString();
                //tracingService.Trace("6");
                payLoad.ModifiedOn = getFormatedDate(contactEntity.GetAttributeValue<DateTime>("modifiedon").AddMinutes(330).ToString(), tracingService);
                //tracingService.Trace("Get All Data");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            return payLoad;
        }
        public static Integration IntegrationConfiguration(IOrganizationService service, string Param, ITracingService tracingService)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            return output;
        }

        public static void CallAPIInsertUpdateAddress(IOrganizationService service, Entity addressEntity, ITracingService tracingService)
        {
            try
            {
                //Guid addressId = new Guid("5de8eb66-7b4e-eb11-a812-0022486e7305");
                //tracingService.Trace("CallAPIInsertUpdateAddress");
                var payload = getAddressDetails(service, addressEntity, tracingService);
                String requestStr = "[" + JsonConvert.SerializeObject(payload) + "]";
                //tracingService.Trace("Trace:- ******************");
                //tracingService.Trace(requestStr);

                Integration integration = IntegrationConfiguration(service, "AddressToMDMSync", tracingService);
                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri;
                ////tracingService.Trace("***********************************");
                ////tracingService.Trace("requestStr " + requestStr);
                ////tracingService.Trace("***********************************");
                String response = string.Empty;
                using (WebClient webClient = new WebClient())
                {
                    webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + _authInfo;
                    webClient.Headers[HttpRequestHeader.ContentType] = "application/json";
                    response = webClient.UploadString(sUrl, requestStr);
                }

                //tracingService.Trace(response);
                response = response.Replace('[', ' ').Replace(']', ' ');
                var responseAddress = Newtonsoft.Json.JsonConvert.DeserializeObject<InsertUpdateAddressResponse>(response);
                if (responseAddress.ErrorMSG == "")
                {
                    Entity _address = new Entity("hil_address");
                    _address["hil_mdmid"] = responseAddress.MDMId != "" ? responseAddress.MDMId : null;
                    _address["hil_errormsg"] = responseAddress.ErrorMSG != "" ? responseAddress.ErrorMSG : null;
                    _address.Id = addressEntity.Id;
                    service.Update(_address);
                    //tracingService.Trace("Error : " + responseAddress.ErrorMSG);
                }
                //tracingService.Trace("Done");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
        public static InsertUpdateAddressPayLoad getAddressDetails(IOrganizationService service, Entity addressEntity, ITracingService tracingService)
        {
            InsertUpdateAddressPayLoad payLoad = new InsertUpdateAddressPayLoad();
            try
            {
                // Entity addressEntity = service.Retrieve("hil_address", addressID, new ColumnSet(true));
                //tracingService.Trace("getAddressDetails");
                payLoad.Name = addressEntity.Contains("hil_name") ? addressEntity.GetAttributeValue<String>("hil_name").ToString() : null;
                //tracingService.Trace("hil_customer");

                payLoad.Street1 = addressEntity.Contains("hil_street1") ? addressEntity.GetAttributeValue<String>("hil_street1").ToString() : null;
                //tracingService.Trace("hil_street1");
                payLoad.Street2 = addressEntity.Contains("hil_street2") ? addressEntity.GetAttributeValue<string>("hil_street2").ToString() : null;
                //tracingService.Trace("hil_street2");
                payLoad.Street3 = addressEntity.Contains("hil_street3") ? addressEntity.GetAttributeValue<string>("hil_street3").ToString() : null;
                //tracingService.Trace("hil_street3");

                payLoad.BusinessGeoMapping = addressEntity.Contains("hil_businessgeo") ? addressEntity.GetAttributeValue<EntityReference>("hil_businessgeo").Id.ToString() : null;
                payLoad.BusinessGeoMappingName = addressEntity.Contains("hil_businessgeo") ? addressEntity.GetAttributeValue<EntityReference>("hil_businessgeo").Name.ToString() : null;
                //tracingService.Trace("hil_businessgeo");

                payLoad.PinCode = addressEntity.Contains("hil_pincode") ? addressEntity.GetAttributeValue<EntityReference>("hil_pincode").Id.ToString() : null;
                payLoad.PinCodeName = addressEntity.Contains("hil_pincode") ? addressEntity.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : null;
                //tracingService.Trace("hil_pincode");

                payLoad.Area = addressEntity.Contains("hil_area") ? addressEntity.GetAttributeValue<EntityReference>("hil_area").Id.ToString() : null;
                payLoad.AreaName = addressEntity.Contains("hil_area") ? addressEntity.GetAttributeValue<EntityReference>("hil_area").Name.ToString() : null;
                //tracingService.Trace("hil_area");

                payLoad.AddressType = addressEntity.Contains("hil_addresstype") ? addressEntity.GetAttributeValue<OptionSetValue>("hil_addresstype").Value.ToString() : null;
                payLoad.AddressTypeName = addressEntity.Contains("hil_addresstype") ? addressEntity.FormattedValues["hil_addresstype"].ToString() : null;
                //tracingService.Trace("hil_addresstype");

                payLoad.Branch = addressEntity.Contains("hil_branch") ? addressEntity.GetAttributeValue<EntityReference>("hil_branch").Id.ToString() : null;
                payLoad.BranchName = addressEntity.Contains("hil_branch") ? addressEntity.GetAttributeValue<EntityReference>("hil_branch").Name.ToString() : null;
                //tracingService.Trace("hil_branch");

                payLoad.City = addressEntity.Contains("hil_city") ? addressEntity.GetAttributeValue<EntityReference>("hil_city").Id.ToString() : null;
                payLoad.CityName = addressEntity.Contains("hil_city") ? addressEntity.GetAttributeValue<EntityReference>("hil_city").Name.ToString() : null;
                //tracingService.Trace("hil_city");

                payLoad.Customer = addressEntity.Contains("hil_customer") ? addressEntity.GetAttributeValue<EntityReference>("hil_customer").Id.ToString() : null;
                payLoad.CustomerName = addressEntity.Contains("hil_customer") ? addressEntity.GetAttributeValue<EntityReference>("hil_customer").Name.ToString() : null;
                //tracingService.Trace("hil_customer");

                payLoad.District = addressEntity.Contains("hil_district") ? addressEntity.GetAttributeValue<EntityReference>("hil_district").Id.ToString() : null;
                payLoad.DistrictName = addressEntity.Contains("hil_district") ? addressEntity.GetAttributeValue<EntityReference>("hil_district").Name.ToString() : null;
                //tracingService.Trace("hil_district");

                payLoad.FullAddress = addressEntity.Contains("hil_fulladdress") ? addressEntity.GetAttributeValue<String>("hil_fulladdress").ToString() : null;
                //tracingService.Trace("hil_fulladdress");

                payLoad.Region = addressEntity.Contains("hil_region") ? addressEntity.GetAttributeValue<EntityReference>("hil_region").Id.ToString() : null;
                payLoad.RegionName = addressEntity.Contains("hil_region") ? addressEntity.GetAttributeValue<EntityReference>("hil_region").Name.ToString() : null;
                //tracingService.Trace("hil_region");

                payLoad.SalesOffice = addressEntity.Contains("hil_salesoffice") ? addressEntity.GetAttributeValue<EntityReference>("hil_salesoffice").Id.ToString() : null;
                payLoad.SalesOfficeName = addressEntity.Contains("hil_salesoffice") ? addressEntity.GetAttributeValue<EntityReference>("hil_salesoffice").Name.ToString() : null;
                //tracingService.Trace("hil_salesoffice");

                payLoad.State = addressEntity.Contains("hil_state") ? addressEntity.GetAttributeValue<EntityReference>("hil_state").Id.ToString() : null;
                payLoad.StateName = addressEntity.Contains("hil_state") ? addressEntity.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : null;
                //tracingService.Trace("hil_state");

                payLoad.SubTerritory = addressEntity.Contains("hil_subterritory") ? addressEntity.GetAttributeValue<EntityReference>("hil_subterritory").Id.ToString() : null;
                payLoad.SubTerritoryName = addressEntity.Contains("hil_subterritory") ? addressEntity.GetAttributeValue<EntityReference>("hil_subterritory").Name.ToString() : null;
                //tracingService.Trace("hil_subterritory");

                payLoad.Latitude = addressEntity.Contains("hil_latitude") ? addressEntity.GetAttributeValue<String>("hil_latitude") : null;
                payLoad.Longitude = addressEntity.Contains("hil_longitude") ? addressEntity.GetAttributeValue<String>("hil_longitude") : null;
                //tracingService.Trace("hil_longitude");

                payLoad.AddressId = addressEntity.Id.ToString();
                //tracingService.Trace("AddressId");

                payLoad.Createdby = addressEntity.Contains("createdby") ? addressEntity.GetAttributeValue<EntityReference>("createdby").Id.ToString() : null;
                payLoad.CreatedbyName = addressEntity.Contains("createdby") ? addressEntity.GetAttributeValue<EntityReference>("createdby").Name.ToString() : null;
                //tracingService.Trace("createdby");

                payLoad.CreatedOn = addressEntity.Contains("createdon") ? getFormatedDate(addressEntity.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString(), tracingService) : null;
                //tracingService.Trace("createdon");

                payLoad.Modifiedby = addressEntity.Contains("modifiedby") ? addressEntity.GetAttributeValue<EntityReference>("modifiedby").Id.ToString() : null;
                payLoad.ModifiedByName = addressEntity.Contains("modifiedby") ? addressEntity.GetAttributeValue<EntityReference>("modifiedby").Name.ToString() : null;
                //tracingService.Trace("modifiedby");

                payLoad.ModifiedOn = addressEntity.Contains("modifiedon") ? getFormatedDate(addressEntity.GetAttributeValue<DateTime>("modifiedon").AddMinutes(330).ToString(), tracingService) : null;
                //tracingService.Trace("modifiedon");

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
            return payLoad;
        }
    }

    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class InsertUpdateCustomerPayLoad
    {
        public string Salutation { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string FullName { get; set; }
        public string EmailAddress { get; set; }
        public string Consent { get; set; }
        public string ConsentName { get; set; }
        public string SourceOfCreation { get; set; }
        public string SourceOfCreationName { get; set; }
        public string MobileNumber { get; set; }
        public string Gender { get; set; }
        public string GenderName { get; set; }
        public string DateOfBirth { get; set; }
        public string DateOfAnniversary { get; set; }
        public string AlternateNumber { get; set; }
        public string ContactId { get; set; }
        public string Createdby { get; set; }
        public string CreatedbyName { get; set; }
        public string CreatedOn { get; set; }
        public string Modifiedby { get; set; }
        public string ModifiedByName { get; set; }
        public string ModifiedOn { get; set; }
    }
    public class InsertUpdateCustomerResponse
    {
        public string MDMId { get; set; }
        public string ContactId { get; set; }
        public string ErrorMSG { get; set; }
    }
    public class InsertUpdateAddressPayLoad
    {
        public string Name { get; set; }
        public string Street1 { get; set; }
        public string Street2 { get; set; }
        public string Street3 { get; set; }
        public string BusinessGeoMapping { get; set; }
        public string BusinessGeoMappingName { get; set; }
        public string PinCode { get; set; }
        public string PinCodeName { get; set; }
        public string Area { get; set; }
        public string AreaName { get; set; }
        public string AddressType { get; set; }
        public string AddressTypeName { get; set; }
        public string Branch { get; set; }
        public string BranchName { get; set; }
        public string City { get; set; }
        public string CityName { get; set; }
        public string Customer { get; set; }
        public string CustomerName { get; set; }
        public string District { get; set; }
        public string DistrictName { get; set; }
        public string FullAddress { get; set; }
        public string Region { get; set; }
        public string RegionName { get; set; }
        public string SalesOffice { get; set; }
        public string SalesOfficeName { get; set; }
        public string State { get; set; }
        public string StateName { get; set; }
        public string SubTerritory { get; set; }
        public string SubTerritoryName { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string AddressId { get; set; }
        public string Createdby { get; set; }
        public string CreatedbyName { get; set; }
        public string CreatedOn { get; set; }
        public string Modifiedby { get; set; }
        public string ModifiedByName { get; set; }
        public string ModifiedOn { get; set; }
    }
    public class InsertUpdateAddressResponse
    {
        public string MDMId { get; set; }
        public string AddressId { get; set; }
        public string ErrorMSG { get; set; }
    }
}
