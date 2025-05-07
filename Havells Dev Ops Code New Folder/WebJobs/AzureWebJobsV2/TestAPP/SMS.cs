using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAPP
{
    public class SMS
    {
        public static void GetSMSList(IOrganizationService service)
        {
            QueryExpression Query = new QueryExpression("hil_smsconfiguration");
            Query.ColumnSet = new ColumnSet("hil_message", "regardingobjectid", "hil_smstemplate");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("regardingobjectid", ConditionOperator.NotNull);
            Query.Criteria.AddCondition("createdon", ConditionOperator.Yesterday);
            Query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
            EntityCollection paymentCollection = service.RetrieveMultiple(Query);

            foreach (Entity sms in paymentCollection.Entities)
            {
                
            }

        }
    }
}
