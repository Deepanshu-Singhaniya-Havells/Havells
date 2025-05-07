using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Deployment;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;
using System.Globalization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ProfileUpdate
    {
        [DataMember(IsRequired = false)]
        public string UName { get; set; }
        [DataMember]
        public string Method { get; set; }
        [DataMember(IsRequired = false)]
        public string Pwd { get; set; }
        [DataMember(IsRequired = false)]
        public int Salutation { get; set; }
        [DataMember(IsRequired = false)]
        public string LName { get; set; }
        [DataMember(IsRequired = false)]
        public string ContGuid { get; set; }//Output
        [DataMember(IsRequired = false)]
        public bool status { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string ErrCode { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string ErrDesc { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string Mob { get; set; }
        [DataMember(IsRequired = false)]
        public string FName { get; set; }
        [DataMember(IsRequired = false)]
        public string Add1 { get; set; }
        [DataMember(IsRequired = false)]
        public string Add2 { get; set; }
        [DataMember(IsRequired = false)]
        public string Add3 { get; set; }
        [DataMember(IsRequired = false)]
        public string ACity { get; set; }
        [DataMember(IsRequired = false)]
        public string AState { get; set; }
        [DataMember(IsRequired = false)]
        public string APinCode { get; set; }
        [DataMember(IsRequired = false)]
        public string Ship1 { get; set; }
        [DataMember(IsRequired = false)]
        public string Ship2 { get; set; }
        [DataMember(IsRequired = false)]
        public string Ship3 { get; set; }
        [DataMember(IsRequired = false)]
        public string SState { get; set; }
        [DataMember(IsRequired = false)]
        public string SCity { get; set; }
        [DataMember(IsRequired = false)]
        public string SPinCode { get; set; }
        [DataMember(IsRequired = false)]
        public string Date_of_Birth { get; set; }
        [DataMember(IsRequired = false)]
        public string Alternate_Mob { get; set; }
        [DataMember(IsRequired = false)]
        public string Date_of_Anniversary { get; set; }

        [DataMember]
        public string ContactGuId { get; set; }

        public ProfileUpdate UpdateProfile(ProfileUpdate Brdg)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            try
            {
                Contact Cont = new Contact();
                Cont.Id = new Guid(Brdg.ContactGuId);
                if (Brdg.FName != null)
                {
                    Cont.FirstName = Brdg.FName;
                }
                if (Brdg.LName != null)
                {
                    Cont.LastName = Brdg.LName;
                }
                if (Brdg.Mob != null)
                {
                    Cont.MobilePhone = Brdg.Mob;
                }
                if (Brdg.Salutation != null)
                {
                    int Salutation = Brdg.Salutation;
                    Cont.hil_Salutation = new OptionSetValue(Salutation);
                }
                if (Brdg.UName != null)
                {
                    Cont.EMailAddress1 = Brdg.UName;
                }
                if (Brdg.Alternate_Mob != null)
                {
                    Cont.Address1_Telephone3 = Brdg.Alternate_Mob;
                }
                if (Brdg.Date_of_Anniversary != null)
                {
                    try
                    {
                        DateTime DOA = DateTime.ParseExact(Brdg.Date_of_Anniversary, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                        Cont["hil_dateofanniversary"] = DOA;
                    }
                    catch
                    {

                    }
                }
                else
                    Cont["hil_dateofanniversary"] = null;
                if (Brdg.Date_of_Birth != null)
                {
                    try
                    {
                        DateTime DOB = DateTime.ParseExact(Brdg.Date_of_Birth, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                        Cont["hil_dateofbirth"] = DOB;
                    }
                    catch
                    {
                    }
                }
                else
                    Cont["hil_dateofbirth"] = null;
                service.Update(Cont);
                Guid PermAddressGuid = GetAddressPermanent(service, Brdg.ContactGuId, 1);
                if (PermAddressGuid != Guid.Empty)
                {
                    hil_address iAddr = new hil_address();
                    iAddr.Id = PermAddressGuid;
                    if (Brdg.Add1 != null)
                        iAddr.hil_Street1 = Brdg.Add1;
                    if (Brdg.Add2 != null)
                        iAddr.hil_Street2 = Brdg.Add2;
                    if (Brdg.Add3 != null)
                        iAddr.hil_Street3 = Brdg.Add3;
                    if (Brdg.ACity != null)
                        iAddr.hil_CIty = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.ACity));
                    if (Brdg.AState != null)
                        iAddr.hil_State = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.AState));
                    if (Brdg.APinCode != null)
                        iAddr.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.APinCode));
                    iAddr.hil_AddressType = new OptionSetValue(1);
                    iAddr.hil_name = "Permanent";
                    service.Update(iAddr);
                }
                else
                {
                    hil_address iAddr = new hil_address();
                    if (Brdg.Add1 != null)
                        iAddr.hil_Street1 = Brdg.Add1;
                    if (Brdg.Add2 != null)
                        iAddr.hil_Street2 = Brdg.Add2;
                    if (Brdg.Add3 != null)
                        iAddr.hil_Street3 = Brdg.Add3;
                    if (Brdg.ACity != null)
                        iAddr.hil_CIty = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.ACity));
                    if (Brdg.AState != null)
                        iAddr.hil_State = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.AState));
                    if (Brdg.APinCode != null)
                        iAddr.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.APinCode));
                    iAddr.hil_Customer = new EntityReference(Contact.EntityLogicalName, new Guid(Brdg.ContactGuId));
                    iAddr.hil_AddressType = new OptionSetValue(1);
                    iAddr.hil_name = "Permanent";
                    //service.Create(iAddr);
                }
                Guid ShippingGuid = GetAddressPermanent(service, Brdg.ContactGuId, 2);
                if (ShippingGuid != Guid.Empty)
                {
                    hil_address iAddr = new hil_address();
                    iAddr.Id = ShippingGuid;
                    if (Brdg.Ship1 != null)
                        iAddr.hil_Street1 = Brdg.Ship1;
                    if (Brdg.Ship2 != null)
                        iAddr.hil_Street2 = Brdg.Ship2;
                    if (Brdg.Ship3 != null)
                        iAddr.hil_Street3 = Brdg.Ship3;
                    if (Brdg.SCity != null)
                        iAddr.hil_CIty = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.SCity));
                    if (Brdg.SState != null)
                        iAddr.hil_State = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.SState));
                    if (Brdg.SPinCode != null)
                        iAddr.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.SPinCode));
                    iAddr.hil_AddressType = new OptionSetValue(2);
                    iAddr.hil_name = "Alternate";
                    service.Update(iAddr);
                }
                else
                {
                    hil_address iAddr = new hil_address();
                    if (Brdg.Ship1 != null)
                        iAddr.hil_Street1 = Brdg.Ship1;
                    if (Brdg.Ship2 != null)
                        iAddr.hil_Street2 = Brdg.Ship2;
                    if (Brdg.Ship3 != null)
                        iAddr.hil_Street3 = Brdg.Ship3;
                    if (Brdg.SCity != null)
                        iAddr.hil_CIty = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.SCity));
                    if (Brdg.SState != null)
                        iAddr.hil_State = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.SState));
                    if (Brdg.SPinCode != null)
                        iAddr.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.SPinCode));
                    iAddr.hil_Customer = new EntityReference(Contact.EntityLogicalName, new Guid(Brdg.ContactGuId));
                    iAddr.hil_AddressType = new OptionSetValue(2);
                    iAddr.hil_name = "Alternate";
                    //service.Create(iAddr);
                }
                Brdg.status = true;
                Brdg.ErrCode = "SUCCESS";
                Brdg.ErrDesc = "";
            }
            catch (Exception ex)
            {
                Brdg.status = false;
                Brdg.ErrCode = "FAILURE";
                Brdg.ErrDesc = ex.Message.ToUpper();
            }
            return (Brdg);
        }
        public static Guid GetAddressPermanent(IOrganizationService service, string iContGuid, int AddType)
        {
            Guid iAddId = new Guid();
            iAddId = Guid.Empty;
            QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, new Guid(iContGuid));
            Query.Criteria.AddCondition("hil_addresstype", ConditionOperator.Equal, AddType);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iAddId = Found.Entities[0].Id;
            }
            return iAddId;
        }
    }
}