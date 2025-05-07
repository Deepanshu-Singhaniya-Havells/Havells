using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace OABooking
{
    public class Models
    {
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
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }

    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class ITTABLEHEADER
    {
        public string ZSHIP_PARTY { get; set; }
        public string ZORDERTYPE { get; set; }
        public string ZTENDERNO { get; set; }
        public string ZNAMEOFCLIENT { get; set; }
        public string ZPROJNAME { get; set; }
        public string ZPO_FOOTER { get; set; }
        public string ZPO_LOINO { get; set; }
        public string ZPO_DATE { get; set; }
        public string ZPRICES { get; set; }
        public string ZPAYTERM { get; set; }
        public string ZMANUFCT_CLEAR { get; set; }
        public string ZUSAGE { get; set; }
        public string ZDATA_GTP { get; set; }
        public string ZQAP { get; set; }
        public string ZINSPECTION { get; set; }
        public string ZZLEADCOD { get; set; }
        public string ZHEADER_TEXT { get; set; }
        public string ZSALES_REP { get; set; }
        public string ZBRANCH_PH { get; set; }
        public string ZZONAL_HEAD { get; set; }
        public string ZBUSINES_UH { get; set; }
    }
    public class ITTABLEITEM
    {
        public string ZTENDERNO { get; set; }
        public string ZTENDER_LINE { get; set; }
        public string ZMATNR { get; set; }
        public string ZQUANTITY { get; set; }
        public string ZUOM { get; set; }
        public string ZWERKS { get; set; }
        public string ZNET_PRICE { get; set; }
        public string ZZFRE { get; set; }
        public string ZLD_PENALTY { get; set; }
        public string ZINDIVIDUAL { get; set; }
        public string ZOVERALL { get; set; }
        public string ZAPP_PRICE { get; set; }
        public string ZDELIVERY_DATE { get; set; }

    }
    public class RootCheckList
    {
        public List<ITTABLEHEADER> IT_TABLE_HEADER { get; set; }
        public List<ITTABLEITEM> IT_TABLE_ITEM { get; set; }
    }
    public class ETTABLE
    {
        public int? VKBUR { get; set; }
        public string ZTENDERNO { get; set; }
        public string ZTENDER_LINE { get; set; }
        public string ZDATA_GTP { get; set; }
        public string ORDERNUMBER { get; set; }
        public int? POSNR { get; set; }
        public string ORDERCREATION { get; set; }
        public string CUST_CODE { get; set; }
        public string CUST_NAME { get; set; }
        public string SHIP_CODE { get; set; }
        public string SHIP_NAME { get; set; }
        public string ITEM_CODE { get; set; }
        public string ITEM_DESC { get; set; }
        public string QTY { get; set; }
        public string MATKL { get; set; }
        public string VALUE { get; set; }
        public string DELIVERY_DATE { get; set; }
        public string ZLD_PENALTY { get; set; }
        public string STOCK { get; set; }
        public string CATEGORY { get; set; }
        public string CATEGORY_DISC { get; set; }
        public string CREDITDAY_CUST { get; set; }
        public string CREDITDAYS { get; set; }
        public string CREDITLIMIT { get; set; }
        public string HEADER_TEXT { get; set; }
        public string REMARKS { get; set; }
    }
    public class RootCheckListReturn
    {
        public string RETURN { get; set; }
        public List<ETTABLE> ET_TABLE { get; set; }
    }
    public class oaheader
    {
        public string hil_oaheaderid { get; set; }
    }
    public class lstOAHeader
    {
        public List<oaheader> oaheaders { get; set; }
    }

    public class LTTABLE
    {
        public string ZTENDERNO { get; set; }
        public string ZTENDER_LINE { get; set; }
        public string READDATE1 { get; set; }
        public string READDATE2 { get; set; }
        public string READDATE3 { get; set; }
        public string INSP_DATE { get; set; }
        public string MODIFY_DATE { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }
    public class OutputClass
    {
        public string EV_RETURN { get; set; }
        public List<LTTABLE> ET_TABLE { get; set; }
    }
}
