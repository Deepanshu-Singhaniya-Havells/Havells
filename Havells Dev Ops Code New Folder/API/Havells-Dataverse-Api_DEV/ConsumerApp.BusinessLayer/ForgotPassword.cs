//ERROR CODE : 1 - USERNAME NOT FOUND

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
    public class ForgotPassword
    {
        [DataMember]
        public string UName { get; set; }
        [DataMember(IsRequired = false)]
        public string ContGuid { get; set; }
        [DataMember(IsRequired = false)]
        public bool IfValid { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }
        [DataMember(IsRequired = false)]
        public int ERROR_CODE { get; set; }
       
        public ForgotPassword RunAction(ForgotPassword bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection

            QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.Or);
            Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, bridge.UName);
            Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, bridge.UName);
            EntityCollection Found1 = service.RetrieveMultiple(Query);
            if (Found1.Entities.Count > 0)
            {
                Contact Cont = (Contact)Found1.Entities[0];
                bridge.ContGuid = Cont.ContactId.Value.ToString();
                bridge.ERROR_CODE = 0;
                bridge.IfValid = true;
                bridge.status = true;
            }
            else
            {
                bridge.ContGuid = "";
                bridge.ERROR_CODE = 1;
                bridge.IfValid = false;
                bridge.status = false;
            }
            return (bridge);
        }
    }
}
