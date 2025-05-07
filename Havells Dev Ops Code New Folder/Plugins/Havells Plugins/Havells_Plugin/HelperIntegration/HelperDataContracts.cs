using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells_Plugin.HelperIntegration
{

    //Serial number validation
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class SerialNumberValidation
    {

        [System.Runtime.Serialization.DataMemberAttribute()]
        public EX_PRD_DET EX_PRD_DET;
    }

    // Type created for JSON at <<root>> --> EX_PRD_DET
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class EX_PRD_DET
    {

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SERIAL_NO;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string MATNR;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string MAKTX;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string SPART;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string REGIO;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string VBELN;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string FKDAT;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string KUNAG;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string NAME1;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string WTY_STATUS;
    }
    

    //Serial number ChildAsset
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class ChildAsset
    {

        [System.Runtime.Serialization.DataMemberAttribute()]
        public ET_SERIAL_DETAIL[] ET_SERIAL_DETAIL;
    }

    // Type created for JSON at <<root>> --> ET_SERIAL_DETAIL
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class ET_SERIAL_DETAIL
    {

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string PARENT_LABEL_ID;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string PARENT_MATNR;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string PART_LABEL_ID;

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string PART_MATNR;
    }
    public class PartnerRootObject
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
    }

    public class BIZGeoMapModelClass
    {
        public string dm_pin_efdt { get; set; }
        public string DM_AREA_EFDT { get; set; }
        public string dm_city_EFDT { get; set; }
        public string dm_dist_EFDT { get; set; }
        public string DM_SALES_OFFICE_EFDT { get; set; }
        public string DM_BRANCH_EFDT { get; set; }
        public string dm_region_EFDT { get; set; }
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_area_desc { get; set; }
        public string sap_city_desc { get; set; }
        public string sap_dist_desc { get; set; }
        public string SAP_SALES_OFFICE_DESC { get; set; }
        public string dm_branch_desc { get; set; }
        public string sap_region_desc { get; set; }
        public string sap_state_desc { get; set; }
        public string DM_SUB_TER_DESC { get; set; }
        public object dm_sub_ter_efdt { get; set; }
        public object dm_state_efdt { get; set; }
        public string dm_pin { get; set; }
        public string DM_AREA { get; set; }
        public string dm_city { get; set; }
        public string dm_dist { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string DM_BRANCH { get; set; }
        public string dm_region { get; set; }
        public string dm_state { get; set; }
        public string dm_sub_ter { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

}
