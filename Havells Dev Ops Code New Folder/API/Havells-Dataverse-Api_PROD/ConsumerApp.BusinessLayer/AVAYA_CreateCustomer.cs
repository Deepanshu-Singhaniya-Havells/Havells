using System;
using System.Net;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Deployment;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    public class ReturnInfo
    {
        [DataMember(IsRequired = false)]
        public string ErrorCode { get; set; }

        [DataMember(IsRequired = false)]
        public string ErrorDescription { get; set; }

        [DataMember(IsRequired = false)]
        public Guid CustomerGuid { get; set; }
    }
    public class AVAYA_CreateCustomer
    {
        [DataMember(IsRequired = false)]

        public string FirstName { get; set; }

        [DataMember(IsRequired = false)]

        public string LastName { get; set; }

        [DataMember(IsRequired = false)]

        public string Mobile { get; set; }

        [DataMember(IsRequired = false)]

        public string Email { get; set; }

        [DataMember(IsRequired = false)]

        public string PermAddress1 { get; set; }

        [DataMember(IsRequired = false)]

        public string PermAddress2 { get; set; }

        [DataMember(IsRequired = false)]

        public string PermLandmark { get; set; }

        [DataMember(IsRequired = false)]

        public string PermCity { get; set; }

        [DataMember(IsRequired = false)]

        public string PermState { get; set; }

        [DataMember(IsRequired = false)]

        public string PermPincode { get; set; }

        [DataMember(IsRequired = false)]

        public string ShipAddress1 { get; set; }

        [DataMember(IsRequired = false)]

        public string ShipAddress2 { get; set; }

        [DataMember(IsRequired = false)]

        public string ShipLandmark { get; set; }

        [DataMember(IsRequired = false)]

        public string ShipCity { get; set; }

        [DataMember(IsRequired = false)]

        public string ShipState { get; set; }

        [DataMember(IsRequired = false)]

        public string ShipPincode { get; set; }

        //[DataMember(IsRequired = false)]

        //public string Brand { get; set; }
        public ReturnInfo CreateCustomer(AVAYA_CreateCustomer Cust)
        {
            ReturnInfo _retCust = new ReturnInfo();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (Cust.Mobile != null && Cust.Email != null)
                {
                    Guid IfAlreadyExists = CheckIfExists(service, Cust.Mobile, Cust.Email);
                    if(IfAlreadyExists == Guid.Empty)
                    {
                        Contact Cont = new Contact();
                        if (Cust.FirstName != null)
                            Cont.FirstName = Cust.FirstName;
                        if (Cust.LastName != null)
                            Cont.LastName = Cust.LastName;
                        Cont.MobilePhone = Cust.Mobile;
                        Cont.EMailAddress1 = Cust.Email;
                        Guid ContId = service.Create(Cont);
                        if (ContId != Guid.Empty)
                        {
                            if (Cust.PermPincode != null)
                            {
                                hil_address Permanent = new hil_address();
                                if (Cust.PermAddress1 != null)
                                    Permanent.hil_Street1 = Cust.PermAddress1.ToUpper();
                                if (Cust.PermAddress2 != null)
                                    Permanent.hil_Street2 = Cust.PermAddress2.ToUpper();
                                if (Cust.PermLandmark != null)
                                    Permanent.hil_Street3 = Cust.PermLandmark.ToUpper();
                                if (Cust.PermCity != null)
                                {
                                    EntityReference City = GetThisCity(service, Cust.PermCity);
                                    if (City != null)
                                        Permanent.hil_CIty = City;
                                }
                                if (Cust.PermPincode != null)
                                {
                                    EntityReference Pincode = GetThisPinCode(service, Cust.PermPincode);
                                    if (Pincode != null)
                                        Permanent.hil_PinCode = Pincode;
                                }
                                if (Cust.PermState != null)
                                {
                                    EntityReference State = GetThisState(service, Cust.PermState);
                                    if (State != null)
                                        Permanent.hil_State = State;
                                }
                                Permanent.hil_AddressType = new OptionSetValue(1);
                                service.Create(Permanent);
                            }
                            if (Cust.ShipPincode != null)
                            {
                                hil_address Shipping = new hil_address();
                                if (Cust.PermAddress1 != null)
                                    Shipping.hil_Street1 = Cust.ShipAddress1.ToUpper();
                                if (Cust.PermAddress2 != null)
                                    Shipping.hil_Street2 = Cust.ShipAddress2.ToUpper();
                                if (Cust.PermLandmark != null)
                                    Shipping.hil_Street3 = Cust.ShipLandmark.ToUpper();
                                if (Cust.ShipCity != null)
                                {
                                    EntityReference City = GetThisCity(service, Cust.ShipCity);
                                    if (City != null)
                                        Shipping.hil_CIty = City;
                                }
                                if (Cust.ShipPincode != null)
                                {
                                    EntityReference Pincode = GetThisPinCode(service, Cust.ShipPincode);
                                    if (Pincode != null)
                                        Shipping.hil_PinCode = Pincode;
                                }
                                if (Cust.ShipState != null)
                                {
                                    EntityReference State = GetThisState(service, Cust.ShipState);
                                    if (State != null)
                                        Shipping.hil_State = State;
                                }
                                Shipping.hil_AddressType = new OptionSetValue(2);
                                service.Create(Shipping);
                                _retCust.CustomerGuid = ContId;
                                _retCust.ErrorCode = "SUCCESS";
                                _retCust.ErrorDescription = "";
                            }
                        }
                    }
                    else
                    {
                        _retCust.CustomerGuid = IfAlreadyExists;
                        _retCust.ErrorCode = "SUCCESS";
                        _retCust.ErrorDescription = "";
                    }
                }
                else
                {
                    _retCust.CustomerGuid = Guid.Empty;
                    _retCust.ErrorCode = "FAILURE";
                    _retCust.ErrorDescription = "EMAIL & MOBILE ARE MANDATORY";
                }
            }
            catch (Exception ex)
            {
                _retCust.CustomerGuid = Guid.Empty;
                _retCust.ErrorCode = "FAILURE";
                _retCust.ErrorDescription = ex.Message.ToUpper();
            }

            return _retCust;

        }
        public static EntityReference GetThisCity(IOrganizationService service, string City)
        {
            EntityReference CityRef = new EntityReference();
            QueryExpression Query = new QueryExpression(hil_city.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, City);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                CityRef = new EntityReference(hil_city.EntityLogicalName, Found.Entities[0].Id);
            }
            return CityRef;
        }
        public static EntityReference GetThisState(IOrganizationService service, string State)
        {
            EntityReference StateRef = new EntityReference();
            QueryExpression Query = new QueryExpression(hil_state.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, State);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                StateRef = new EntityReference(hil_state.EntityLogicalName, Found.Entities[0].Id);
            }
            return StateRef;
        }
        public static EntityReference GetThisPinCode(IOrganizationService service, string PinCode)
        {
            EntityReference PinCodeRef = new EntityReference();
            QueryExpression Query = new QueryExpression(hil_pincode.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, PinCode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                PinCodeRef = new EntityReference(hil_pincode.EntityLogicalName, Found.Entities[0].Id);
            }
            return PinCodeRef;
        }
        public static Guid CheckIfExists(IOrganizationService service, string CustMobile, string CustEmail)
        {
            Guid ContId = new Guid();
            ContId = CheckThisUserName(service, CustEmail);
            if(ContId == Guid.Empty)
            {
                ContId = CheckThisMobileNumber(service, CustMobile);
            }
            return ContId;
        }
        public static Guid CheckThisUserName(IOrganizationService service, string Email)
        {
            Guid ContId = new Guid();
            QueryByAttribute Query = new QueryByAttribute();
            Query.EntityName = Contact.EntityLogicalName;
            Query.AddAttributeValue("emailaddress1", Email);
            ColumnSet Col = new ColumnSet("contactid");
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                ContId = Found.Entities[0].Id;
            }
            return ContId;
        }
        public static Guid CheckThisMobileNumber(IOrganizationService service, string MobNum)
        {
            Guid VD = new Guid();
            QueryByAttribute Query = new QueryByAttribute();
            Query.EntityName = Contact.EntityLogicalName;
            Query.AddAttributeValue("mobilephone", MobNum);
            ColumnSet Col = new ColumnSet("contactid");
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                VD = Found.Entities[0].Id;
            }
            return VD;
        }
    }
}