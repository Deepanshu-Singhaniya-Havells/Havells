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
using Havells_Plugin;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AVAYACreateWo
    {
        [DataMember]
        public string MOBILE { get; set; }
        [DataMember]
        public string PIN { get; set; }
        [DataMember]
        public string CALL_STYPE { get; set; }
        [DataMember]
        public string DIV { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public AVAYACreateWo AvayaCreateWrkOrder(AVAYACreateWo Cust)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                Guid _thisCust = GetThisCustomer(service, Cust.MOBILE);
                if (_thisCust != Guid.Empty)
                {
                    Guid GetDivision = GetGuidByName2(Product.EntityLogicalName, "name", Cust.DIV, service);
                    Guid ServAcc = GetGuidByName2(Account.EntityLogicalName, "name", "Dummy Customer", service);
                    Guid PriceList = GetGuidByName2(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                    Guid PinCode = GetGuidByName2(hil_pincode.EntityLogicalName, "hil_name", Cust.PIN, service);
                    Guid CallSubType = GetGuidByName2(hil_callsubtype.EntityLogicalName, "hil_name", Cust.CALL_STYPE, service);
                    msdyn_workorder Job = new msdyn_workorder();
                    Job.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, _thisCust);
                    if (CallSubType != Guid.Empty)
                    {
                        Job.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, CallSubType);
                        hil_callsubtype Call = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, CallSubType, new ColumnSet("hil_calltype"));
                        if (Call.hil_CallType != null)
                        {
                            Job.hil_CallType = Call.hil_CallType;
                        }
                    }
                    if (GetDivision != Guid.Empty)
                        Job.hil_Productcategory = new EntityReference(Product.EntityLogicalName, GetDivision);
                    if (ServAcc != Guid.Empty)
                    {
                        Job.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServAcc);
                        Job.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServAcc);
                    }
                    if (PriceList != Guid.Empty)
                        Job.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
                    if (PinCode != Guid.Empty)
                        Job.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, PinCode);
                    Guid JobId = service.Create(Job);
                    if (JobId != Guid.Empty)
                    {
                        msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("msdyn_name"));
                        Cust.RESULT = "SUCCESS";
                        Cust.RESULT_DESC = enJob.msdyn_name;
                    }
                    else
                    {
                        Cust.RESULT = "FAILURE";
                        Cust.RESULT_DESC = "TEMPORARY CRM FAILURE, TRY AGAIN LATER";
                    }
                }
                else
                {
                    Cust.RESULT = "FAILURE";
                    Cust.RESULT_DESC = "CUSTOMER NOT FOUND";
                }
            }
            catch (Exception ex)
            {
                Cust.RESULT = "FAILURE";
                Cust.RESULT_DESC = ex.Message.ToUpper();
            }
            return (Cust);
        }
        public static Guid GetThisCustomer(IOrganizationService service, string Mobile)
        {
            Guid _thisCust = new Guid();
            QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, Mobile));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                _thisCust = Found.Entities[0].Id;
            }
            return _thisCust;
        }
        public static Guid GetGuidByName2(string _eName, string _fName, string _qName, IOrganizationService service)
        {
            Guid ThisGuid = new Guid();
            QueryExpression Query = new QueryExpression(_eName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression(_fName, ConditionOperator.Equal, _qName));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                ThisGuid = Found.Entities[0].Id;
            }
            return ThisGuid;
        }
    }
}