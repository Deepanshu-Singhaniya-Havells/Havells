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
    public class GetCustomerGuid
    {
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember(IsRequired = false)]
        public string ContactId { get; set; }
        public GetCustomerGuid GetGuIdBasisMobNo(GetCustomerGuid Cust)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "contact";
            ColumnSet Col = new ColumnSet("fullname");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, Cust.MobileNumber));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            if (Colec.Entities.Count == 1)
            {
                foreach (Entity et in Colec.Entities)
                {
                    Cust.ContactId = Convert.ToString(et.Id);
                }
            }
            else
            {
                Cust.ContactId = "Doesn't Exist";
            }
            return (Cust);
        }
    }
}