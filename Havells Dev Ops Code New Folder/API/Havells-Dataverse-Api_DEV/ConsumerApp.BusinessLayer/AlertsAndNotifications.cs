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
    public class AlertsAndNotifications
    {
        [DataMember]
        public string QueryDate { get; set; }
        [DataMember(IsRequired = false)]
        public string Message { get; set; }
        public List<AlertsAndNotifications> GetAlertsOnDate(AlertsAndNotifications Alrt)
        {
            DateTime QryDate = Convert.ToDateTime(Alrt.QueryDate);
            List<AlertsAndNotifications> obj = new List<AlertsAndNotifications>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "hil_alerts";
            ColumnSet Col = new ColumnSet("hil_message", "hil_publishdate");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("hil_publishdate", ConditionOperator.LessEqual, QryDate));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_publishdate", ConditionOperator.GreaterEqual, QryDate.AddMonths(-1)));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (Entity et in Colec.Entities)
            {
                obj.Add(
                new AlertsAndNotifications
                {
                    Message = Convert.ToString(et["hil_message"])
                });
            }
            return (obj);
        }
    }
}