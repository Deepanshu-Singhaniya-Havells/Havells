using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InventoryWebJobs
{
    public class IntegrationLog
    {
        public Guid hil_integrationlogid { get; set; }
        public string hil_name { get; set; }
        public int hil_totalrecords { get; set; }
        public int hil_recordsaffected { get; set; }
        public int hil_totalerrorrecords { get; set; }
        public bool hil_states { get; set; }
        public string hil_syncdate { get; set; }
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


}