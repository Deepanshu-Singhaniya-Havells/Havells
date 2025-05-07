using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
namespace Havells_Plugin.AccountEn
{
    public class Common
    {

        public static void SetFullAddressAcc(Entity entity, String sMessage, IOrganizationService service)
        {
            try
            {
                if (sMessage == "CREATE")
                {
                    //tracingService.Trace("Common1");
                    Account enAccount = entity.ToEntity<Account>();
                    hil_address MyAddress = new hil_address();

                    String sFullAddress = String.Empty;

                    Account upAccount = new Account();
                    upAccount.AccountId = entity.Id;

                    if (enAccount.Address1_Line1 != null)
                    {
                        //tracingService.Trace("Common2");
                        sFullAddress = enAccount.Address1_Line1 + " ";
                        MyAddress.hil_Street1 = enAccount.Address1_Line1;
                    }
                    if (enAccount.Address1_Line1 != null)
                    {

                        MyAddress.hil_Street2 = enAccount.Address1_Line2;
                        //tracingService.Trace("Common3");
                        sFullAddress += enAccount.Address1_Line2 + " ";
                    }
                    if (enAccount.Address1_Line3 != null)
                    {
                        MyAddress.hil_Street3 = enAccount.Address1_Line3;
                        //tracingService.Trace("Common4");
                        sFullAddress += enAccount.Address1_Line3 + " ";
                    }
                    if (enAccount.hil_city != null)
                    {

                        MyAddress.hil_CIty = enAccount.hil_city;
                        upAccount.hil_TextCity = enAccount.hil_city.Name;
                        //tracingService.Trace("Common5" + enAccount.hil_city.Name);
                        //tracingService.Trace("Common5" + enAccount.hil_city.Id);
                        sFullAddress += enAccount.hil_city.Name + " ";
                    }
                    if (enAccount.hil_district != null)
                    {
                        MyAddress.hil_District = enAccount.hil_district;
                        upAccount.hil_TextDistrict = enAccount.hil_district.Name;
                        //tracingService.Trace("Common6" + enAccount.hil_district.Name);
                        //tracingService.Trace("Common6" + enAccount.hil_district.Id);
                        sFullAddress += enAccount.hil_district.Name + " ";
                    }
                    if (enAccount.hil_pincode != null)
                    {
                        MyAddress.hil_PinCode = enAccount.hil_pincode;
                        upAccount.hil_TextPinCode = enAccount.hil_pincode.Name;
                        //tracingService.Trace("Common7" + enAccount.hil_pincode.Name);
                        //tracingService.Trace("Common7" + enAccount.hil_pincode.Id);
                        sFullAddress += enAccount.hil_pincode.Name + " ";
                    }
                    else return;
                    if (enAccount.hil_area != null)
                    {
                        //MyAddress.hil_a= enAccount.hil_area;
                        upAccount.hil_TextArea = enAccount.hil_area.Name;
                        //tracingService.Trace("Common8" + enAccount.hil_area.Name);
                        //tracingService.Trace("Common8" + enAccount.hil_area.Id);
                        sFullAddress += enAccount.hil_area.Name + " ";
                    }
                    if (enAccount.hil_state != null)
                    {
                        MyAddress.hil_State = enAccount.hil_state;
                        upAccount.hil_TextState = enAccount.hil_state.Name;
                        //tracingService.Trace("Common9" + enAccount.hil_state.Name);
                        //tracingService.Trace("Common9" + enAccount.hil_state.Id);
                        sFullAddress += enAccount.hil_state.Name + " ";
                    }
                    upAccount.hil_FullAddress = sFullAddress;
                    service.Update(upAccount);


                    MyAddress.hil_FullAddress = sFullAddress;
                    MyAddress.hil_AddressType = new OptionSetValue(1);
                    MyAddress.hil_name = "PermanentAddress";
                    MyAddress.hil_Customer = new EntityReference(Account.EntityLogicalName, upAccount.Id);
                    service.Create(MyAddress);
                }
                else if (sMessage == "UPDATE")
                {
                    //tracingService.Trace("Common1");
                    Account enAccount = entity.ToEntity<Account>();
                    hil_address MyAddress = new hil_address();

                    String sFullAddress = String.Empty;

                    Account upAccount = new Account();
                    upAccount.AccountId = entity.Id;

                    Guid fsPrimaryAddressId = Guid.Empty;
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        var obj = from _Address in orgContext.CreateQuery<hil_address>()
                                  where _Address.hil_AddressType.Value == 1
                                  && _Address.hil_Customer.Id == entity.Id
                                  select new
                                  {
                                      _Address.hil_addressId
                                  ,
                                      _Address.hil_AddressType
                                  };
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_AddressType.Value == 1)
                                fsPrimaryAddressId = iobj.hil_addressId.Value;
                        }
                    }
                    if (enAccount.Address1_Line1 != null)
                    {
                        //tracingService.Trace("Common2");
                        sFullAddress = enAccount.Address1_Line1 + " ";
                        MyAddress.hil_Street1 = enAccount.Address1_Line1;
                    }
                    if (enAccount.Address1_Line1 != null)
                    {

                        MyAddress.hil_Street2 = enAccount.Address1_Line2;
                        //tracingService.Trace("Common3");
                        sFullAddress += enAccount.Address1_Line2 + " ";
                    }
                    if (enAccount.Address1_Line3 != null)
                    {
                        MyAddress.hil_Street3 = enAccount.Address1_Line3;
                        //tracingService.Trace("Common4");
                        sFullAddress += enAccount.Address1_Line3 + " ";
                    }
                    if (enAccount.hil_city != null)
                    {

                        MyAddress.hil_CIty = enAccount.hil_city;
                        upAccount.hil_TextCity = enAccount.hil_city.Name;
                        //tracingService.Trace("Common5" + enAccount.hil_city.Name);
                        //tracingService.Trace("Common5" + enAccount.hil_city.Id);
                        sFullAddress += enAccount.hil_city.Name + " ";
                    }
                    if (enAccount.hil_district != null)
                    {
                        MyAddress.hil_District = enAccount.hil_district;
                        upAccount.hil_TextDistrict = enAccount.hil_district.Name;
                        //tracingService.Trace("Common6" + enAccount.hil_district.Name);
                        //tracingService.Trace("Common6" + enAccount.hil_district.Id);
                        sFullAddress += enAccount.hil_district.Name + " ";
                    }
                    if (enAccount.hil_pincode != null)
                    {
                        MyAddress.hil_PinCode = enAccount.hil_pincode;
                        upAccount.hil_TextPinCode = enAccount.hil_pincode.Name;
                        //tracingService.Trace("Common7" + enAccount.hil_pincode.Name);
                        //tracingService.Trace("Common7" + enAccount.hil_pincode.Id);
                        sFullAddress += enAccount.hil_pincode.Name + " ";
                    }
                    else return;
                    if (enAccount.hil_area != null)
                    {
                        //MyAddress.hil_a= enAccount.hil_area;
                        upAccount.hil_TextArea = enAccount.hil_area.Name;
                        //tracingService.Trace("Common8" + enAccount.hil_area.Name);
                        //tracingService.Trace("Common8" + enAccount.hil_area.Id);
                        sFullAddress += enAccount.hil_area.Name + " ";
                    }
                    if (enAccount.hil_state != null)
                    {
                        MyAddress.hil_State = enAccount.hil_state;
                        upAccount.hil_TextState = enAccount.hil_state.Name;
                        //tracingService.Trace("Common9" + enAccount.hil_state.Name);
                        //tracingService.Trace("Common9" + enAccount.hil_state.Id);
                        sFullAddress += enAccount.hil_state.Name + " ";
                    }
                    upAccount.hil_FullAddress = sFullAddress;
                    service.Update(upAccount);


                    MyAddress.hil_FullAddress = sFullAddress;
                    MyAddress.hil_AddressType = new OptionSetValue(1);
                    if (fsPrimaryAddressId != Guid.Empty)
                    {
                        MyAddress.hil_addressId = fsPrimaryAddressId;
                        service.Update(MyAddress);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ContactEn.SetFullAddress" + ex.Message);
            }
        }


    }
}
