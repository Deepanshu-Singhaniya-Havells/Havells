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
    public class CallSubtypeMaster
    {
        [DataMember(IsRequired = false)]
        public string CallSubTypeName { get; set; }
        [DataMember(IsRequired = false)]
        public string CallSubTypeId { get; set; }
        public List<CallSubtypeMaster> GetAllCallSubType()
        {
            List<CallSubtypeMaster> obj = new List<CallSubtypeMaster>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "hil_callsubtype";
            ColumnSet Col = new ColumnSet("hil_name", "hil_callsubtypeid");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.Or);
            //Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Demo"));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Breakdown"));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Installation"));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Pre Sale Demo"));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (Entity et in Colec.Entities)
            {
                obj.Add(
                new CallSubtypeMaster
                {
                    CallSubTypeName = Convert.ToString(et["hil_name"]),
                    CallSubTypeId = Convert.ToString(et["hil_callsubtypeid"])
                });
            }
            return (obj);
        }
    }
}
