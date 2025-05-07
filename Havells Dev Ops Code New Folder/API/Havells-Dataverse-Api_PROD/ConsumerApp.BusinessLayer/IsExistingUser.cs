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

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IsExistingUser
    {
        [DataMember]
        public Int32 SOURCE { get; set; }
        [DataMember]
        public string VALUE { get; set; }
        [DataMember(IsRequired = false)]
        public bool IfExisting { get; set; }
        public IsExistingUser CheckIfDuplicate(IsExistingUser iCheck)
        {
            //IOrganizationService service = ConnectToCRM.GetOrgService();
            //try
            //{
            //    QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            //    Query.ColumnSet = new ColumnSet(false);
            //    Query.Criteria = new FilterExpression(LogicalOperator.Or);
            //    if(iCheck.SOURCE == 1)
            //    {
            //        Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, iCheck.VALUE);
            //    }
            //    else if(iCheck.SOURCE == 2)
            //    {
            //        Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, iCheck.VALUE);
            //    }
            //    EntityCollection Found = service.RetrieveMultiple(Query);
            //    if (Found.Entities.Count > 0)
            //    {
            //        iCheck.IfExisting = true;
            //    }
            //    else
            //    {
            //        iCheck.IfExisting = false;
            //    }
            //}
            //catch(Exception ex)
            //{

            //}
            iCheck.IfExisting = true;
            return iCheck;
        }
    }
}
