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
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ProfileView
    {
        [DataMember]
        public string ContactGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string Email { get; set; }
        [DataMember(IsRequired = false)]
        public string Salutation { get;  set; }
        [DataMember(IsRequired = false)]
        public string LName { get; set; }
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
        public string ACountry { get; set; }
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
        public string SCountry { get; set; }
        [DataMember(IsRequired = false)]
        public string SCity { get; set; }
        [DataMember(IsRequired = false)]
        public string SPinCode { get; set; }
        [DataMember(IsRequired = false)]
        public DateTime Date_of_Birth { get; set; }
        [DataMember(IsRequired = false)]
        public string Alternate_Mob { get; set; }
        [DataMember(IsRequired = false)]
        public DateTime Date_of_Anniversary { get; set; }
        public ProfileView GetCustomerInformation(ProfileView Pf)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (Pf.ContactGuId != null)
                {
                    Contact Contact1 = (Contact)service.Retrieve(Contact.EntityLogicalName, new Guid(Pf.ContactGuId), new ColumnSet("emailaddress1", "mobilephone",
                        "firstname", "lastname", "hil_salutation", "hil_dateofanniversary", "hil_dateofbirth", "address1_telephone3"));
                    if (Contact1.EMailAddress1 != null)
                        Pf.Email = Contact1.EMailAddress1;
                    if (Contact1.MobilePhone != null)
                        Pf.Mob = Contact1.MobilePhone;
                    if (Contact1.FirstName != null)
                        Pf.FName = Contact1.FirstName;
                    if (Contact1.LastName != null)
                        Pf.LName = Contact1.LastName;
                    if (Contact1.hil_Salutation != null)
                    {
                        OptionSetValue Salutation1 = Contact1.hil_Salutation;
                        Pf.Salutation = Convert.ToString(Salutation1.Value);
                    }
                    else
                    {
                        Pf.Salutation = "0";
                    }
                    if(Contact1.Attributes.Contains("hil_dateofbirth"))
                    {
                        DateTime DOB = (DateTime)Contact1["hil_dateofbirth"];
                        Pf.Date_of_Birth = DOB.AddDays(1);
                    }
                    if (Contact1.Attributes.Contains("hil_dateofanniversary"))
                    {
                        DateTime AnvDate = (DateTime)Contact1["hil_dateofanniversary"];
                        Pf.Date_of_Anniversary = AnvDate.AddDays(1);
                    }
                    if(Contact1.Address1_Telephone3 != null)
                    {
                        Pf.Alternate_Mob = Contact1.Address1_Telephone3;
                    }
                    hil_address PermAddress = FindAddress(service, Pf.ContactGuId, 1);
                    if (PermAddress != null)
                    {
                        if (PermAddress.hil_Street1 != null)
                        { 
                            Pf.Add1 = PermAddress.hil_Street1;
                        }
                        else
                        {
                            Pf.Add1 = " ";
                        }
                        if (PermAddress.hil_Street2 != null)
                        {
                            Pf.Add2 = PermAddress.hil_Street2;
                        }
                        else
                        {
                            Pf.Add2 = " ";
                        }
                        if (PermAddress.hil_Street3 != null)
                        {
                            Pf.Add3 = PermAddress.hil_Street3;
                        }
                        else
                        {
                            Pf.Add3 = " ";
                        }
                        if (PermAddress.hil_CIty != null)
                        {
                            Pf.ACity = PermAddress.hil_CIty.Name;
                        }
                        else
                        {
                            Pf.ACity = " ";
                        }
                        if (PermAddress.hil_State != null)
                        {
                            Pf.AState = PermAddress.hil_State.Name;
                        }
                        else
                        {
                            Pf.AState = " ";
                        }
                        if (PermAddress.hil_PinCode != null)
                        {
                            Pf.APinCode = PermAddress.hil_PinCode.Name;
                        }
                        else
                        {
                            Pf.APinCode = " ";
                        }
                    }
                    else
                    {
                        Pf.Add1 = " ";
                        Pf.Add2 = " ";
                        Pf.Add3 = " ";
                        Pf.ACity = " ";
                        Pf.AState = " ";
                        Pf.APinCode = " ";
                    }
                    hil_address ShipAddress = FindAddress(service, Pf.ContactGuId, 2);
                    if (ShipAddress != null)
                    {
                        if (ShipAddress.hil_Street1 != null)
                        {
                            Pf.Ship1 = ShipAddress.hil_Street1;
                        }
                        else
                        {
                            Pf.Ship1 = " ";
                        }
                        if (ShipAddress.hil_Street2 != null)
                        {
                            Pf.Ship2 = ShipAddress.hil_Street2;
                        }
                        else
                        {
                            Pf.Ship2 = " ";
                        }
                        if (ShipAddress.hil_Street3 != null)
                        {
                            Pf.Ship3 = ShipAddress.hil_Street3;
                        } 
                        else
                        {
                            Pf.Ship3 = " ";
                        }
                        if (ShipAddress.hil_CIty != null)
                        {
                            Pf.SCity = ShipAddress.hil_CIty.Name;
                        }
                        else
                        {
                            Pf.SCity = " ";
                        }
                        if (ShipAddress.hil_State != null)
                        {
                            Pf.SState = ShipAddress.hil_State.Name;
                        }
                        else
                        {
                            Pf.SState = " ";
                        }
                        if (ShipAddress.hil_PinCode != null)
                        {
                            Pf.SPinCode = ShipAddress.hil_PinCode.Name;
                        }
                        else
                        {
                            Pf.SPinCode = " ";
                        }
                    }
                    else
                    {
                        Pf.Ship1 = " ";
                        Pf.Ship2 = " ";
                        Pf.Ship3 = " ";
                        Pf.SCity = " ";
                        Pf.SState = " ";
                        Pf.SPinCode = " ";
                    }
                }
            }
            catch(Exception ex)
            {
                Pf.ContactGuId = ex.Message.ToUpper();
            }
            return (Pf);
        }
        public static hil_address FindAddress(IOrganizationService service, string iContGuid, int AddType)
        {
            hil_address iAddId = new hil_address();
            iAddId.Id = Guid.Empty;
            QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, new Guid(iContGuid));
            Query.Criteria.AddCondition("hil_addresstype", ConditionOperator.Equal, AddType);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iAddId = Found.Entities[0].ToEntity<hil_address>();
            }
            return iAddId;
        }
    }
}
