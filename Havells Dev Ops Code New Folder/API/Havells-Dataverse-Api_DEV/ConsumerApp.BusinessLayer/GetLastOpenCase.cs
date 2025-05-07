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
    public class GetLastOpenCase
    {
        [DataMember]
        public string MOBILE { get; set; }
        [DataMember(IsRequired = false)]
        public string STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string VISIT_DATE { get; set; }
        [DataMember(IsRequired = false)]
        public string TECHNICIAN { get; set; }
        [DataMember(IsRequired = false)]
        public string TICKET_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public GetLastOpenCase GetLastOpenCaseBasisMobNo(GetLastOpenCase Cust)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                Guid _thisCust = GetThisCustomer(service, Cust.MOBILE);
                if (_thisCust != Guid.Empty)
                {
                    QueryExpression Query = new QueryExpression("msdyn_workorder");
                    Query.ColumnSet = new ColumnSet("msdyn_name", "msdyn_substatus");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, _thisCust));
                    Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    Query.AddOrder("createdon", OrderType.Descending);
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        Entity Job = (Entity)Found.Entities[0];
                        QueryExpression Qry = new QueryExpression("phonecall");
                        Qry.ColumnSet = new ColumnSet("scheduledend", "ownerid");
                        Qry.Criteria = new FilterExpression(LogicalOperator.And);
                        Qry.Criteria.AddCondition(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, Job.Id));
                        //Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        Qry.AddOrder("createdon", OrderType.Descending);
                        EntityCollection Fnd = service.RetrieveMultiple(Qry);
                        if (Fnd.Entities.Count > 0)
                        {
                            Entity _fPh = (Entity)Fnd.Entities[0];
                            if (_fPh.Attributes.Contains("scheduledend"))
                            {
                                DateTime Due = (DateTime)_fPh["scheduledend"];
                                Cust.VISIT_DATE = Convert.ToString(Due);
                            }
                            else
                            {
                                Cust.VISIT_DATE = "";
                            }
                            EntityReference Owner = (EntityReference)_fPh["ownerid"];
                            Cust.TECHNICIAN = Owner.Name;
                            Cust.TICKET_ID = (string)Job["msdyn_name"];
                            EntityReference Sta = (EntityReference)Job["msdyn_substatus"];
                            Cust.STATUS = Sta.Name;
                            Cust.RESULT = "SUCCESS";
                            Cust.RESULT_DESC = "";
                        }
                        else
                        {
                            Cust.VISIT_DATE = "";
                            Cust.TECHNICIAN = "";
                            Cust.TICKET_ID = (string)Job["msdyn_name"];
                            EntityReference Sta = (EntityReference)Job["msdyn_substatus"];
                            Cust.STATUS = Sta.Name;
                            Cust.RESULT = "SUCCESS";
                            Cust.RESULT_DESC = "NO VISIT PLANNED";
                        }
                    }
                    else
                    {
                        Cust.VISIT_DATE = "";
                        Cust.STATUS = "";
                        Cust.TECHNICIAN = "";
                        Cust.TICKET_ID = "";
                        Cust.RESULT = "SUCCESS";
                        Cust.RESULT_DESC = "NO OPEN CASES FOUND";
                    }
                }
                else
                {
                    Cust.VISIT_DATE = "";
                    Cust.STATUS = "";
                    Cust.TECHNICIAN = "";
                    Cust.TICKET_ID = "";
                    Cust.RESULT = "FAILURE";
                    Cust.RESULT_DESC = "CUSTOMER NOT FOUND";
                }
                return (Cust);
            }
            catch(Exception ex)
            {
                Cust.VISIT_DATE = "";
                Cust.STATUS = "";
                Cust.TECHNICIAN = "";
                Cust.TICKET_ID = "";
                Cust.RESULT = "FAILURE";
                Cust.RESULT_DESC = (ex.Message).ToUpper();
                return (Cust);
            }
        }
        public static Guid GetThisCustomer(IOrganizationService service, string Mobile)
        {
           Guid _thisCust = new Guid();
            QueryExpression Query = new QueryExpression("contact");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, Mobile));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                _thisCust = Found.Entities[0].Id;
            }
            return _thisCust;
        }
    }
}