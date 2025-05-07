using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChannelPartnerSyncJob.Models
{
    public class PartnerResult
    {
        public string sap_city_desc { get; set; }
        public string sap_dist_desc { get; set; }
        public string sap_state_desc { get; set; }
        public string dm_area_desc { get; set; }
        public string KUNNR { get; set; }
        public string KTOKD { get; set; }
        public string PAR_CODE { get; set; }
        public string KTONR { get; set; }
        public string VTXTM { get; set; }
        public string STREET { get; set; }
        public string STR_SUPPL3 { get; set; }
        public string ADDRESS3 { get; set; }
        public string dm_state { get; set; }
        public string dm_dist { get; set; }
        public string dm_city { get; set; }
        public string dm_pin { get; set; }
        public string KUKLA { get; set; }
        public string KNRZE { get; set; }
        public string NAME1 { get; set; }
        public string J_1IPANNO { get; set; }
        public string J_1ILSTNO { get; set; }
        public string BANKA { get; set; }
        public string BANKN { get; set; }
        public string BANKL { get; set; }
        public string SMTP_ADDR { get; set; }
        public string MOB_NUMBER { get; set; }
        public string PAR_NAME { get; set; }
        public string CONT_CODE { get; set; }
        public string CONT_PERSON { get; set; }
        public string STATUS { get; set; }
        public string DM_AREA { get; set; }
        public string SALES_STATUS { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
        public string RPMKR { get; set; }
        public string SAMPARK_MOB_NO { get; set; }
        public string FIRM_TYPE { get; set; }
        public string DM_REGION { get; set; }
        public string DM_BRANCH { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string AADHAAR_NO { get; set; }
        public string GST_NO { get; set; }
        public string MNC_NAME { get; set; }
        public string MNC_MOB_NO { get; set; }
        public string MNC_MAIL_ID { get; set; }
        public string KAM_ID { get; set; }
        public string BUSAB { get; set; }
        public string Longitude { get; set; }
        public string Latitude { get; set; }

    }
    public class PartnerRootObject
    {
        public object Result { get; set; }
        public List<PartnerResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class PartnerDivisionResult
    {
        public string KUNNR { get; set; }
        public string dm_ch { get; set; }
        public string dm_div { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string ERDAT { get; set; }
        public string STATUS { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string MODIFYBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MTIMESTAMP { get; set; }
        public string SAP_CH { get; set; }
        public string SAP_DIV { get; set; }
    }
    public class PartnerDivisionRootObject
    {
        public object Result { get; set; }
        public List<PartnerDivisionResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
}
