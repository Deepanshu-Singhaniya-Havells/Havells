using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.WebJob.MasterSync.Helper
{
    public class iHelper
    {
        public static IRestResponse exesuteAPI(string _APIUrl, Dictionary<string, string> header, Dictionary<string, string> parameter, RestSharp.Method method)
        {
            Console.WriteLine("API Call Start");
            Console.WriteLine("API URL: " + _APIUrl);
            var client = new RestClient(_APIUrl);
            client.Timeout = -1;
            var request = new RestRequest(method);
            foreach (KeyValuePair<string, string> keyValue in header)
            {
                request.AddHeader(keyValue.Key, keyValue.Value);
            }
            foreach (KeyValuePair<string, string> keyValue in parameter)
            {
                request.AddParameter(keyValue.Key, keyValue.Value);
            }
            IRestResponse response = client.Execute(request);
            Console.WriteLine("Response: " + response.Content);
            return response;
        }
        public static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                if (output.uri.Contains("middleware.havells.com"))
                {
                    output.uri = output.uri.Replace("middleware", "p90ci");
                }
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        public static DateTime? StringToDateTime(string _mdmTimeStamp)
        {
            Console.WriteLine("StringToDateTime Function Started ");
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            Console.WriteLine("timestamp  " + _dtMDMTimeStamp);
            Console.WriteLine("StringToDateTime Function Ended");
            return _dtMDMTimeStamp;
        }
        public static string DateTimeToString(DateTime _cTimeStamp)
        {
            Console.WriteLine("DateTimeToString Function Started ");
            string timestamp = null;
            timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
            Console.WriteLine("timestamp  " + timestamp);
            Console.WriteLine("DateTimeToString Function Ended");
            return timestamp;
        }
        public static Entity retriveData(string entityName, string fieldName, string fieldsValue, IOrganizationService service)
        {
            Entity _retObj = null;
            QueryExpression Query = new QueryExpression(entityName);
            Query.ColumnSet = new ColumnSet("statecode");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(fieldName, ConditionOperator.Equal, fieldsValue);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                _retObj = Found.Entities[0];
            }
            return _retObj;
        }

    }
  
}
