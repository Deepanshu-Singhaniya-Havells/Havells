using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;

namespace Havells.CRM.DynamicsPOSyncToSAP
{
    public class HelperClass
    {
        public static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service, string APIName)
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
        public static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

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
    public enum WarrantyStatus
    {
        In_Warranty = 1,
        Out_Warranty = 2,
        Warranty_Void = 3,
        NA_for_Warranty = 4
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }
    public class IT_DATA
    {
        public string MATNR { get; set; }
        public int DZMENG { get; set; }
        public string SPART { get; set; }
    }
    public class Parent
    {
        public string AUART { get; set; }
        public string VTWEG { get; set; }
        public string SPART { get; set; }
        public string BSTKD { get; set; }
        public string BSTDK { get; set; }
        public string KUNNR { get; set; }
        public string ABRVW { get; set; }
        public string VKORG { get; set; }
    }
    public class PO_SAP_Request
    {
        public string IM_PROJECT { get; set; }
        public List<IT_DATA> LT_LINE_ITEM { get; set; }
        public Parent IM_HEADER { get; set; }
    }
    public class TRETURN
    {
        public string KUNNR { get; set; }
        public string EBELN { get; set; }
        public string VBELN { get; set; }
        public string MESSAGE { get; set; }
    }
    public class IMPROJECT
    {
        public string IM_PROJECT { get; set; }
    }
    public class RootObject
    {
        public string EX_SALESDOC_NO { get; set; }
        public string RETURN { get; set; }
        public List<TRETURN> T_RETURN { get; set; }
    }
    public class PO_SAP_Response
    {
        public Response ZBAPI_CREATE_SALES_ORDER;
    }
    public class Response
    {
        public string EX_SALESDOC_NO { get; set; }
        public string RETURN { get; set; }
        public List<ET_Details> ET_SO_DETAILS;
    }
    public class ET_Details
    {
        public string VBELN { get; set; }
        public string POSNR { get; set; }
        public string MATNR { get; set; }
        public string MAKTX { get; set; }
        public string ERDAT { get; set; }
        public string KWMENG { get; set; }
        public string VRKME { get; set; }
        public string NETWR { get; set; }
        public string WAERK { get; set; }
        public string NETPR { get; set; }
        public string MWSBP { get; set; }
        public string H_LFSTK { get; set; }
        public string H_STATUS { get; set; }
        public string L_LFSTA { get; set; }
        public string L_STATUS { get; set; }
        public string KUNNR { get; set; }
        public string H_DEL_DATE { get; set; }
        public string L_DEL_DATE { get; set; }
        public string BSTNK { get; set; }
        public string IHREZ { get; set; }
        public string H_NETWR { get; set; }
        public string H_WAERK { get; set; }
        public string ST_KUNNR { get; set; }
        public string NAME1 { get; set; }
        public string NAME2 { get; set; }
        public string LAND1 { get; set; }
        public string ORT01 { get; set; }
        public string PSTLZ { get; set; }
        public string REGIO { get; set; }
        public string STRAS { get; set; }
        public string TELF1 { get; set; }
        public string TELFX { get; set; }
        public string REGSMTP_ADDRIO { get; set; }
    }

    public class LTINVOICE
    {
        public string VBELN { get; set; }
        public string POSNR { get; set; }
        public string FKART { get; set; }
        public string FKDAT { get; set; }
        public string AUBEL { get; set; }
        public string NETWR { get; set; }
        public string MWSBK { get; set; }
        public string TOT_DISC { get; set; }
        public string KUNRG { get; set; }
        public int WERKS { get; set; }
        public string BUTXT { get; set; }
        public string BEZEI { get; set; }
        public string MATNR { get; set; }
        public string FKIMG { get; set; }
        public string MEINS { get; set; }
        public string CHARG { get; set; }
        public string MRP { get; set; }
        public string DLP { get; set; }
        public string PUR_RATE { get; set; }
        public string LLAMOUNT { get; set; }
        public string LLDISCOUNT { get; set; }
        public string LLTAXAMT { get; set; }
        public string LLTOTAMT { get; set; }
        public string BT_MANUF_DATE { get; set; }
        public string BT_EXP_DATE { get; set; }
        public string ROUNDOFF { get; set; }
        public string FKSTO { get; set; }
        public string TAXABLE_AMT { get; set; }
        public string CGST_PERC { get; set; }
        public string SGST_PERC { get; set; }
        public string UTGST_PERC { get; set; }
        public string IGST_PERC { get; set; }
        public string CESS_PERC { get; set; }
        public string CGST_AMT { get; set; }
        public string SGST_AMT { get; set; }
        public string UTGST_AMT { get; set; }
        public string IGST_AMT { get; set; }
        public int STEUC { get; set; }
        public string REL_PARTY { get; set; }
        public string INV_TYPE { get; set; }
    }
    public class InvoiceResonse
    {
        public List<LTINVOICE> LT_INVOICE { get; set; }
    }
}
