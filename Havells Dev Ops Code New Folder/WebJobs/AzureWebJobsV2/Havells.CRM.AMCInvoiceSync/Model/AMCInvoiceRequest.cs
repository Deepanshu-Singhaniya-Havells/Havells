using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.AMCInvoiceSync.Model
{
    public class AMCInvoiceRequest
    {
        public List<AMCInvoiceData> IM_DATA { get; set; }
    }

    public class AMCInvoiceData
    {
        public string START_DATE { get; set; }
        public string END_DATE { get; set; }
        public string SOLD_TO_PARTY { get; set; }
        public string SHIP_TO_PARTY { get; set; }
        public string SHIP_TO_PA_NAME { get; set; }
        public string STREET { get; set; }
        public string TITLE { get; set; }
        public string CUSTOMER_NAME { get; set; }
        public string NAME2 { get; set; }
        public string HOUSE_NO { get; set; }
        public string STREET4 { get; set; }
        public string STREET5 { get; set; }
        public string POSTAL_CODE { get; set; }
        public string DISTRICT { get; set; }
        public string CITY { get; set; }
        public string REGION_CODE { get; set; }
        public string CALL_ID { get; set; }
        public string MATERIAL_CODE { get; set; }
        public string CS_SERVICE_PRICE { get; set; }
        public string PRODUCT_SERIAL_NO { get; set; }
        public string PRODUCT_DOP { get; set; }
        public string EMPLOYEE_ID { get; set; }
        public string EMP_TYPE { get; set; }
        public string EMAIL { get; set; }
        public string PHONE { get; set; }
        public string REMARKS { get; set; }
        public string BRANCH_NAME { get; set; }
        public string STATUS { get; set; }
        public string WARNTY_TILL_DATE { get; set; }
        public string WARNTY_STATUS { get; set; }
        public string CLOSED_ON { get; set; }
        public string PRICING_DATE { get; set; }
        public string ZDS6_AMOUNT { get; set; }
        public string DISCOUNT_FLAG { get; set; }
        public string PAYMENT_REF_NO { get; set; }
    }
}