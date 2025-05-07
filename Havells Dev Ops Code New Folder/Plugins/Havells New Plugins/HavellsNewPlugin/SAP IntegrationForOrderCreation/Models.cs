using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.SAP_IntegrationForOrderCreation
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
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            return output;
        }
    }
    // Root myDeserializedClass = JsonConvert.Deserializestring<Root>(myJsonResponse);
    public class RespnseCreditCheck
    {
        public string CRD_STATUS { get; set; }
        public string CRD_TEXT { get; set; }
        public string CRD_LIMIT_AMT { get; set; }
        public string CRD_LIMIT_PERC { get; set; }
        public string RECEIVABLE { get; set; }
        public string LIABILITY { get; set; }
        public string SALES_VALUE { get; set; }
        public string CREDIT_EXPO { get; set; }
        public string CREDIT_DAYS { get; set; }
    }
    public class RootCreditCheck
    {
        public string I_KKBER { get; set; }
        public string I_KUNNR { get; set; }
        public string I_VKORG { get; set; }
        public string I_REGUL { get; set; }
    }
    public class RootGetOpenItems
    {
        public string CUSTOMER { get; set; }
        public string KEYDATE { get; set; }
    }
    public class RETURN
    {
        public string TYPE { get; set; }
        public string CODE { get; set; }
        public string MESSAGE { get; set; }
        public string LOG_NO { get; set; }
        public string LOG_MSG_NO { get; set; }
        public string MESSAGE_V1 { get; set; }
        public string MESSAGE_V2 { get; set; }
        public string MESSAGE_V3 { get; set; }
        public string MESSAGE_V4 { get; set; }
    }
    public class LINEITEM
    {
        public string COMP_CODE { get; set; }
        public string CUSTOMER { get; set; }
        public string SP_GL_IND { get; set; }
        public string CLEAR_DATE { get; set; }
        public string CLR_DOC_NO { get; set; }
        public string ALLOC_NMBR { get; set; }
        public string FISC_YEAR { get; set; }
        public string DOC_NO { get; set; }
        public string ITEM_NUM { get; set; }
        public string PSTNG_DATE { get; set; }
        public string DOC_DATE { get; set; }
        public string ENTRY_DATE { get; set; }
        public string CURRENCY { get; set; }
        public string LOC_CURRCY { get; set; }
        public string REF_DOC_NO { get; set; }
        public string DOC_TYPE { get; set; }
        public string FIS_PERIOD { get; set; }
        public string POST_KEY { get; set; }
        public string DB_CR_IND { get; set; }
        public string BUS_AREA { get; set; }
        public string TAX_CODE { get; set; }
        public string LC_AMOUNT { get; set; }
        public string AMT_DOCCUR { get; set; }
        public string LC_TAX { get; set; }
        public string TX_DOC_CUR { get; set; }
        public string ITEM_TEXT { get; set; }
        public string BRANCH { get; set; }
        public string BLINE_DATE { get; set; }
        public string PMNTTRMS { get; set; }
        public string DSCT_DAYS1 { get; set; }
        public string DSCT_DAYS2 { get; set; }
        public string NETTERMS { get; set; }
        public string DSCT_PCT1 { get; set; }
        public string DSCT_PCT2 { get; set; }
        public string DISC_BASE { get; set; }
        public string DSC_AMT_LC { get; set; }
        public string DSC_AMT_DC { get; set; }
        public string PYMT_METH { get; set; }
        public string PMNT_BLOCK { get; set; }
        public string FIXEDTERMS { get; set; }
        public string INV_REF { get; set; }
        public string INV_YEAR { get; set; }
        public string INV_ITEM { get; set; }
        public string DUNN_BLOCK { get; set; }
        public string DUNN_KEY { get; set; }
        public string LAST_DUNN { get; set; }
        public string DUNN_LEVEL { get; set; }
        public string DUNN_AREA { get; set; }
        public string DOC_STATUS { get; set; }
        public string NXT_DOCTYP { get; set; }
        public string VAT_REG_NO { get; set; }
        public string REASON_CDE { get; set; }
        public string PMTMTHSUPL { get; set; }
        public string REF_KEY_1 { get; set; }
        public string REF_KEY_2 { get; set; }
        public string T_CURRENCY { get; set; }
        public string AMOUNT { get; set; }
        public string NET_AMOUNT { get; set; }
        public string NAME { get; set; }
        public string NAME_2 { get; set; }
        public string NAME_3 { get; set; }
        public string NAME_4 { get; set; }
        public string POSTL_CODE { get; set; }
        public string CITY { get; set; }
        public string COUNTRY { get; set; }
        public string STREET { get; set; }
        public string PO_BOX { get; set; }
        public string POBX_PCD { get; set; }
        public string POBK_CURAC { get; set; }
        public string BANK_ACCT { get; set; }
        public string BANK_KEY { get; set; }
        public string BANK_CTRY { get; set; }
        public string TAX_NO_1 { get; set; }
        public string TAX_NO_2 { get; set; }
        public string TAX { get; set; }
        public string EQUAL_TAX { get; set; }
        public string REGION { get; set; }
        public string CTRL_KEY { get; set; }
        public string INSTR_KEY { get; set; }
        public string PAYEE_CODE { get; set; }
        public string LANGU { get; set; }
        public string BILL_LIFE { get; set; }
        public string BE_TAXCODE { get; set; }
        public string BILLTAX_LC { get; set; }
        public string BILLTAX_FC { get; set; }
        public string LC_COL_CHG { get; set; }
        public string COLL_CHARG { get; set; }
        public string CHGS_TX_CD { get; set; }
        public string ISSUE_DATE { get; set; }
        public string USAGEDATE { get; set; }
        public string BILL_USAGE { get; set; }
        public string DOMICILE { get; set; }
        public string DRAWER { get; set; }
        public string CTRBNK_LOC { get; set; }
        public string DRAW_CITY1 { get; set; }
        public string DRAWEE { get; set; }
        public string DRAW_CITY2 { get; set; }
        public string DISCT_DAYS { get; set; }
        public string DISCT_RATE { get; set; }
        public string ACCEPTED { get; set; }
        public string BILLSTATUS { get; set; }
        public string PRTEST_IND { get; set; }
        public string BE_DEMAND { get; set; }
        public string OBJ_TYPE { get; set; }
        public string REF_DOC { get; set; }
        public string REF_ORG_UN { get; set; }
        public string REVERSAL_DOC { get; set; }
        public string SP_GL_TYPE { get; set; }
        public string NEG_POSTNG { get; set; }
        public string REF_DOC_NO_LONG { get; set; }
        public string BILL_DOC { get; set; }
    }
    public class TITEM
    {
        public string BLART { get; set; }
        public string BELNR { get; set; }
        public string XBLNR { get; set; }
        public string BUDAT { get; set; }
        public string DMBTRO { get; set; }
        public string ZFBDT { get; set; }
        public string NODAY { get; set; }
        public string SP_GL_IND { get; set; }
        public string NETTERMS { get; set; }
        public string BLINE_DATE { get; set; }
        public string DB_CR_IND { get; set; }
        public string DOC_TYPE { get; set; }
        public string ITEM_TEXT { get; set; }
        public string DOC_NO { get; set; }
        public string AMT_DOCCUR { get; set; }
        public string DOC_DATE { get; set; }
        public string CLEAR_DATE { get; set; }
    }
    public class CCODE
    {
        public int BUKRS { get; set; }
    }
    public class RespnseGetOpenItems
    {
        public RETURN RETURN { get; set; }
        public string OVERDUE { get; set; }
        public string NET_DUE { get; set; }
        public List<LINEITEM> LINEITEMS { get; set; }
        public List<TITEM> T_ITEMS { get; set; }
        public List<CCODE> C_CODES { get; set; }
    }



    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class ITTABLEHEADER
    {
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
}
