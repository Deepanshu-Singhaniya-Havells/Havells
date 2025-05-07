using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Crm.PriceListSync
{
    public class Model
    {
        public static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password", "hil_contactno");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                if (integrationConfiguration.Contains("hil_contactno"))
                {
                    output.MTIMESTAMP = integrationConfiguration.GetAttributeValue<string>("hil_contactno");
                }
                else
                    output.MTIMESTAMP = "19000101000000";
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
        public string MTIMESTAMP { get; set; }
    }
    public class RequestPayload
    {
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public bool IsInitialLoad { get; set; }
        public string Condition { get; set; }

    }
    public class PricingRoot
    {
        public object Result { get; set; }
        public List<PricingResult> Results { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class PricingResult
    {
        public string MATNR { get; set; }
        public decimal KBETR { get; set; }
        public string DATAB { get; set; }
        public string DATBI { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string MODIFYBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MTIMESTAMP { get; set; }
    }
}
