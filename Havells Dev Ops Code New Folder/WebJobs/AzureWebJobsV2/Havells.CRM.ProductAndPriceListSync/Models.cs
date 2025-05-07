using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.ProductAndPriceListSync
{
    public class Models
    {
        public static Guid GetStagingProductPrice(string materialcode, DateTime startDate, DateTime endDate, IOrganizationService service)
        {
            Guid iDivision = new Guid();
            iDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("hil_stagingpricingmapping");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, materialcode);
            Query.Criteria.AddCondition("hil_datestart", ConditionOperator.Equal, startDate);
            Query.Criteria.AddCondition("hil_dateend", ConditionOperator.Equal, endDate);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDivision = Found.Entities[0].Id;
            }
            return iDivision;
        }
        public static Guid GetProduct(string materialcode, IOrganizationService service)
        {
            Guid iDivision = new Guid();
            iDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet("hil_division");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, materialcode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDivision = Found.Entities[0].Id;
            }
            return iDivision;
        }
        public static Guid GetProductDivision(string materialcode, IOrganizationService service)
        {
            Guid iDivision = new Guid();
            iDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet("hil_division");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, materialcode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {

                iDivision = Found.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id;
            }
            return iDivision;
        }
        public static DateTime? StringToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }
        public static Integration GetIntegration(IOrganizationService service, string RecName)
        {
            try
            {
                Integration intConf = new Integration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, RecName);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                intConf.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                intConf.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                intConf.passWord = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return intConf;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
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
        public static string getTimeStamp(IOrganizationService service)
        {
            string _enquiryDatetime = "19000101000000";
            QueryExpression qsCType = new QueryExpression("product");
            qsCType.ColumnSet = new ColumnSet("hil_mdmtimestamp");
            qsCType.NoLock = true;
            qsCType.TopCount = 1;
            qsCType.AddOrder("hil_mdmtimestamp", OrderType.Descending);
            EntityCollection entCol = service.RetrieveMultiple(qsCType);
            if (entCol.Entities.Count > 0)
            {
                DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(300);
                if (_cTimeStamp.Year.ToString().PadLeft(4, '0') != "0001")
                    _enquiryDatetime = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + (_cTimeStamp.Second).ToString().PadLeft(2, '0');
            }
            return _enquiryDatetime;
        }
        public static Guid GetGuidbyNameCommon(String sEntityName, String sFieldName, String sFieldValue,
                    IOrganizationService service, int iStatusCode = 0)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                if (iStatusCode == 2)
                {
                    //qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                    qe.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 2);
                }
                //qe.Criteria.AddCondition("hil_deleteflag", ConditionOperator.NotEqual, 1);
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.Helper.GetGuidbyName" + ex.Message);
            }
            return fsResult;
        }
        public static EntityReference GetHSNCodeRefrence(String HSNCode, IOrganizationService service)
        {
            EntityReference fsResult = null;
            try
            {
                QueryExpression qe = new QueryExpression("hil_hsncode");
                qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, HSNCode);
                qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].ToEntityReference() ;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.Helper.GetHSNCodeRefrence" + ex.Message);
            }
            return fsResult;
        }
        public static DateTime? ConvertToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }
        public static Guid CheckForAMCProduct(string materialcode, IOrganizationService service)
        {
            Guid iProduct = Guid.Empty;
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, materialcode);
            Query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iProduct = Found.Entities[0].Id;
            }
            return iProduct;
        }
    }
    public class LTTABLE
    {
        public string VKORG { get; set; }
        public string MATNR { get; set; }
        public string VKBUR { get; set; }
        public string KSCHL { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public string DATBI { get; set; }
        public string DATAB { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }
    public class OutputClass
    {
        public string EV_RETURN { get; set; }
        public List<LTTABLE> LT_TABLE { get; set; }
    }
    public class Integration
    {
        public string uri { get; set; }
        public string userName { get; set; }
        public string passWord { get; set; }
    }
    public class RootObjectProduct
    {
        public object Result { get; set; }
        public List<Product> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class RequestPayload
    {
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public bool IsInitialLoad { get; set; }
        public string Condition { get; set; }
    }
    public class PricingResult
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public double KBETR { get; set; }
        public string KONWA { get; set; }
        public string VKORG { get; set; }
        public DateTime DATAB { get; set; }
        public DateTime DATBI { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CTIMESTAMP { get; set; }
        public string CreatedBy { get; set; }
        public string MTIMESTAMP { get; set; }
        public string ModifyBy { get; set; }
    }
    public class PricingRoot
    {
        public object Result { get; set; }
        public List<PricingResult> Results { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class Product
    {
        public string sap_div_desc;
        public string ProductLine_EWBEZ;
        public string MATNR;
        public string MVGR1;
        public string MATKL;
        public string dm_Div;
        public string EXTWG;
        public string VKORG;
        public string MAKTX;
        public string WGBEZ;
        public string MVGR1_DESC;
        public string MVGR2;
        public string MVGR2_DESC;
        public string MVGR3;
        public string MVGR3_DESC;
        public string MVGR4;
        public string MVGR4_DESC;
        public string MVGR5;
        public string MVGR5_DESC;
        public string EWBEZ;
        public string NTGEW;
        public string GEWEI;
        public string MHDHB;
        public string EAN11;
        public string STATUS;
        public string SERNP;
        public string MTART;
        public string MSTAE;
        public string DELETE_FLAG;
        public string CREATEDBY;
        public string CTIMESTAMP;
        public string MODIFYBY;
        public string MTIMESTAMP;
        public string STEUC;
        public string SAP_DIV;
        public string KONDM;
    }
}
