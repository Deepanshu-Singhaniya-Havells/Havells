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
    public class StateMaster
    {
        [DataMember(IsRequired = false)]
        public string StateName { get; set; }
        [DataMember(IsRequired = false)]
        public string StateGuId { get; set; }
        public List<StateMaster> GetAllActiveStateCodes()
        {
            List<StateMaster> obj = new List<StateMaster>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "hil_state";
            ColumnSet Col = new ColumnSet("hil_name", "hil_stateid");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (Entity et in Colec.Entities)
            {
                obj.Add(
                new StateMaster
                {
                    StateName = Convert.ToString(et["hil_name"]),
                    StateGuId = Convert.ToString(et["hil_stateid"])
                });
            }
            return (obj);
        }
    }
}
