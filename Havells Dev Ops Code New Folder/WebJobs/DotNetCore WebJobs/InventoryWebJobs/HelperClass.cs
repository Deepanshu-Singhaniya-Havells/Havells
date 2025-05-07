using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace InventoryWebJobs
{
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }

    }
    public static class HelperClass
    {
        public static IntegrationConfiguration GetIntegrationConfiguration(ServiceClient _service, string APIName)
        {
            try
            {
                IntegrationConfiguration inconfig = new IntegrationConfiguration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, APIName);
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return inconfig;
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
        public static string CallAPI(IntegrationConfiguration integrationConfiguration, string Json, String method)
        {
            WebRequest request = WebRequest.Create(integrationConfiguration.url);
            request.Headers[HttpRequestHeader.Authorization] = "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(integrationConfiguration.userName + ":" + integrationConfiguration.password));
            request.Method = method; //"POST";
            if (!string.IsNullOrEmpty(Json))
            {
                byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteArray.Length;
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
            }
            WebResponse response = request.GetResponse();
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            Stream dataStream1 = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream1);
            return reader.ReadToEnd();
        }

    }
}
