using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace Havells.Dataverse.CustomConnector.Inventory
{

    public class PO_Invoicedetails : IPlugin

    {
        public static ITracingService tracingService = null;
        IPluginExecutionContext context;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            string JsonResponse = string.Empty;
            try
            {
                string _IM_VBELN = Convert.ToString(context.InputParameters["IM_VBELN"]);

                if (string.IsNullOrWhiteSpace(_IM_VBELN))
                {
                    LTELOG tELOG = new LTELOG() { ERROR = "IM_VBELN is required" };
                    return;
                }

                Reqbody reqbody = new Reqbody
                {
                    IM_VBELN = _IM_VBELN,
                };

                QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "CRM_SAPtoCRM_InvoiceNew");
                Entity enColl = service.RetrieveMultiple(qe)[0];
                string URL = enColl.GetAttributeValue<string>("hil_url");
                string Auth = enColl.GetAttributeValue<string>("hil_username") + ":" + enColl.GetAttributeValue<string>("hil_password");

                var data = new StringContent(JsonConvert.SerializeObject(reqbody), Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(Auth);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                HttpResponseMessage response = client.PostAsync(URL, data).Result;
                if (response.IsSuccessStatusCode)
                {
                    var result = response.Content.ReadAsStringAsync().Result;
                    //	LTINVOICE apiResponse = JsonConvert.DeserializeObject<LTINVOICE>(result);
                    //	JsonResponse = JsonConvert.SerializeObject(result);
                    context.OutputParameters["data"] = result;
                }
            }
            catch (Exception ex)
            {
                var errorLog = new List<LTELOG>
                {
                    new LTELOG { ERROR = "D365 Internal Server Error : " + ex.Message }
                };
                JsonResponse = JsonConvert.SerializeObject(errorLog);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
    }

    public class LTELOG
    {
        public string KEY1 { get; set; }
        public string KEY2 { get; set; }
        public string KEY3 { get; set; }
        public string ERROR { get; set; }
    }

    public class LTINVOICE
    {
        public string BSTNK { get; set; }
        public int BSTDK1 { get; set; }
        public string ORD_NO { get; set; }
        public int ERDAT1 { get; set; }
        public string DEL_NO { get; set; }
        public int ERDAT2 { get; set; }
        public object BIL_NO { get; set; }
        public int FKDAT { get; set; }
        public string ORD_ITEM { get; set; }
        public string MATNR { get; set; }
        public string ARKTX { get; set; }
        public string KWMENG { get; set; }
        public string LFIMG { get; set; }
        public string FKIMG { get; set; }
        public string REJ_RESN { get; set; }
        public string GENO { get; set; }
        public string KUNNR { get; set; }
        public string BOLNR { get; set; }
        public string TRAID { get; set; }
        public string VSART { get; set; }
    }

    public class Root
    {
        public List<LTINVOICE> LT_INVOICE { get; set; }
        public List<LTELOG> LT_ELOG { get; set; }
    }

    public class Reqbody
    {
        public string IM_VBELN { get; set; }
    }
}

