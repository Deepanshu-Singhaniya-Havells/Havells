using Microsoft.Xrm.Sdk;
using System.Net.Http.Headers;
using System.Net;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Configuration;

public class EasyRewardTest
{
    //public static ITracingService tracingService = null;
    public IOrganizationService service;
    public EasyRewardTest(IOrganizationService _service)
    {
        service  = _service;
    }
    public void Execute()
    {
        #region PluginConfig
        //tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
        //IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
        //IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
        //IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
        #endregion
        //try
        //{
        //if (context.InputParameters.Contains("CustomerGuid") && context.InputParameters["CustomerGuid"] is string
        //        && context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string && context.Depth == 1)
        //if (context.InputParameters.Contains("CustomerGuid") && context.InputParameters["CustomerGuid"] is string && context.Depth == 1)
        //{
        string MobileNumber = "";
        //string CustomerGuid = (string)context.InputParameters["CustomerGuid"];
        string CustomerGuid = "74817b3a-2f7d-ee11-8178-6045bdc66add";
        if (!string.IsNullOrEmpty(CustomerGuid))
        {
            MobileNumber = GetMobileNumber(new Guid(CustomerGuid));
            string URL = GetProfileUrl(MobileNumber);
            Console.WriteLine(URL);
            //if (URL != "" || URL != null)
            //{
            //    context.OutputParameters["ReturnMessage"] = "Success";
            //    context.OutputParameters["WidgetURL"] = URL;
            //}
            //else
            //{
            //    context.OutputParameters["ReturnMessage"] = "Failed";
            //    context.OutputParameters["WidgetURL"] = "";
            //}
        }
        //else
        //{
        //    context.OutputParameters["ReturnMessage"] = "Failed";
        //    context.OutputParameters["WidgetURL"] = "";
        //}
        //}
        //}
        //catch (Exception ex)
        //{
        //    //context.OutputParameters["ReturnMessage"] = "D365 Internal Server Error : " + ex.Message;
        //    //context.OutputParameters["WidgetURL"] = "";
        //}
    }
    internal string GetMobileNumber(Guid customerGuid)
    {
        string MobileNumber = "";

        //QueryExpression Query = new QueryExpression("contact");
        //Query.ColumnSet = new ColumnSet("customerguid", "mobilephone");
        //Query.Criteria.AddCondition("customerguid", ConditionOperator.Equal, customerGuid);
        //EntityCollection entColl = service.RetrieveMultiple(Query);
        Entity data = service.Retrieve("contact", customerGuid, new ColumnSet("mobilephone"));
        MobileNumber = data.GetAttributeValue<string>("mobilephone");

        return MobileNumber;
    }
    internal string GetProfileUrl(string mobileNumber)
    {
        IntegrationConfiguration inconfig = GetIntegrationConfiguration("Work flow Widget");//->// To pass the name
        WidgetRequest objWidgetRequest = new WidgetRequest { WidgetCode = "VC" };
        var Paramdata = new StringContent(JsonConvert.SerializeObject(objWidgetRequest), System.Text.Encoding.UTF8, "application/json");
        var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);

        var client = new HttpClient();
        client.DefaultRequestHeaders.Add("Authorization", new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray)).ToString());
        client.DefaultRequestHeaders.Add("LoginUserId", mobileNumber);
        client.DefaultRequestHeaders.Add("OperationToken", "");
        HttpResponseMessage Result = client.PostAsync(inconfig.url, Paramdata).Result;
        if (Result.IsSuccessStatusCode)
        {
            WidgetResponse obj = JsonConvert.DeserializeObject<WidgetResponse>(Result.Content.ReadAsStringAsync().Result);
            if (!string.IsNullOrEmpty(obj.WidgetURL))
            {
                return obj.WidgetURL;
            }
        }
        return "";
    }

    private IntegrationConfiguration GetIntegrationConfiguration(string name)
    {
        try
        {
            IntegrationConfiguration inconfig = new IntegrationConfiguration();
            QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
            qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
            qsCType.NoLock = true;
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, name);
            Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
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
}



public class WidgetRequest
{
    public string WidgetCode { get; set; }
}
public class WidgetResponse
{
    public string WidgetURL { get; set; }
    public string ReturnMessage { get; set; }
    public string ReturnCode { get; set; }
}
public class IntegrationConfiguration
{
    public string url { get; set; }
    public string userName { get; set; }
    public string password { get; set; }
}