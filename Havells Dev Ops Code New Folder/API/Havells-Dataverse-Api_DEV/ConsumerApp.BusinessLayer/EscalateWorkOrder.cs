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
    public class EscalateWorkOrder
    {
        [DataMember]
        public string MOBILE { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public EscalateWorkOrder EscallateWorkOd(EscalateWorkOrder Cust)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            Guid _thisCust = GetThisCustomer(service, Cust.MOBILE);
            if (_thisCust != Guid.Empty)
            {
                QueryExpression Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("msdyn_name", "msdyn_substatus");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, _thisCust));
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Query.AddOrder("createdon", OrderType.Descending);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    msdyn_workorder Job = (msdyn_workorder)Found.Entities[0];
                    Task _crTask = new Task();
                    _crTask.hil_TaskType = new OptionSetValue(2);
                    _crTask.RegardingObjectId = new EntityReference(msdyn_workorder.EntityLogicalName, Job.Id);
                    _crTask.Subject = "CREATED FROM AVAYA";
                    Guid TskId = service.Create(_crTask);
                    if(TskId != Guid.Empty)
                    {
                        Cust.RESULT = "SUCCESS";
                        Cust.RESULT_DESC = "";
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
                    Cust.RESULT_DESC = "NO OPEN CASES FOUND";
                }
            }
            else
            {
                Cust.RESULT = "FAILURE";
                Cust.RESULT_DESC = "CUSTOMER NOT FOUND";
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
    }
}