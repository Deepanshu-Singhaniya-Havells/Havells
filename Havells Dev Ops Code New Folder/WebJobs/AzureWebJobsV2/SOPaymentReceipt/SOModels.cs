using System.Collections.Generic;
using Microsoft.Xrm.Sdk;

namespace SOPaymentReceipt
{
	public class TransactionDetails
	{
		public string transactionID { get; set; }
		public EntityReference Regarding { get; set; }
	}
	public class IntegrationConfig
	{
		public string uri { get; set; }
		public string Auth { get; set; }
	}
	public class StatusRequest
	{

		public string PROJECT { get; set; }

		public string command { get; set; }

		public string var1 { get; set; }

	}
	public class StatusResponse
	{

		public int status { get; set; }

		public string msg { get; set; }

		public List<TransactionDetail> transaction_details { get; set; }
	}
	public class TransactionDetail
	{

		public string mihpayid { get; set; }

		public string request_id { get; set; }

		public string bank_ref_num { get; set; }

		public string amt { get; set; }

		public string transaction_amount { get; set; }

		public string txnid { get; set; }

		public string additional_charges { get; set; }

		public string productinfo { get; set; }

		public string firstname { get; set; }

		public string bankcode { get; set; }

		public string udf1 { get; set; }

		public string udf3 { get; set; }

		public string udf4 { get; set; }

		public string udf5 { get; set; }

		public string field2 { get; set; }

		public string field9 { get; set; }

		public string error_code { get; set; }

		public string addedon { get; set; }

		public string payment_source { get; set; }

		public string card_type { get; set; }

		public string error_Message { get; set; }

		public string net_amount_debit { get; set; }

		public string disc { get; set; }

		public string mode { get; set; }

		public string PG_TYPE { get; set; }

		public string card_no { get; set; }

		public string udf2 { get; set; }

		public string status { get; set; }

		public string unmappedstatus { get; set; }

		public string Merchant_UTR { get; set; }

		public string Settled_At { get; set; }
	}
	public class AMCInvoiceRequest
	{
		public List<AMCInvoiceData> LT_TABLE { get; set; }
	}
	public class AMCInvoiceData
	{
		public string CALL_ID { get; set; }
		public string SOLD_TO_PARTY { get; set; }
		public string SHIP_TO_PARTY { get; set; }
		public string SHIP_TO_PA_NAME { get; set; }
		public string TITLE { get; set; }
		public string CUSTOMER_NAME { get; set; }
		public string NAME_2 { get; set; }
		public string STREET { get; set; }
		public string HOUSE_NO { get; set; }
		public string STREET4 { get; set; }
		public string STREET5 { get; set; }
		public string POSTAL_CODE { get; set; }
		public string DISTRICT { get; set; }
		public string CITY { get; set; }
		public string REGION_CODE { get; set; }
		public string PO_DATE { get; set; }
		public string MATERIAL_CODE { get; set; }
		public string PLANT { get; set; }
		public string CS_SERVICE_PRICE { get; set; }
		public string PRODUCT_SERIAL_NO { get; set; }
		public string PRODUCT_DOP { get; set; }
		public string EMPLOYEE_ID { get; set; }
		public string EMAIL { get; set; }
		public string PHONE { get; set; }
		public string REMARKS { get; set; }
		public string PAYMENT_REF_NO { get; set; }
		public string BRANCH_NAME { get; set; }
		public string ORDER_NUMBER { get; set; }
		public string BILLING_NUMBER { get; set; }
		public string WARNTY_TILL_DATE { get; set; }
		public string WARNTY_STATUS { get; set; }
		public string CLOSED_ON { get; set; }
		public string PRICING_DATE { get; set; }
		public string ZDS6_AMOUNT { get; set; }
		public string DISCOUNT_FLAG { get; set; }
		public string CREATEDBY { get; set; }
		public string CTIMESTAMP { get; set; }
		public string MODIFYBY { get; set; }
		public string MTIMESTAMP { get; set; }
		public string MESSAGE { get; set; }
		public string START_DATE { get; set; }
		public string END_DATE { get; set; }
		public string EMP_TYPE { get; set; }
		public string STATUS { get; set; }
		public string ITEM_NO { get; set; }

	}
	public class ActionResponse
	{
		public bool Is_Successful { get; set; }
		public string Message { get; set; }
	}
	public class AmcReceiptData
	{
		public string ReceiptNumber { get; set; }
		public string CustomerName { get; set; }
		public string ContactNumber { get; set; }
		public string Email { get; set; }
		public string CRMRefNo { get; set; }
		public string TypeOfAMC { get; set; }
		public string Date { get; set; }
		public string Amount { get; set; }
		public string AMCPlan { get; set; }
        public string PaymentMode { get; set; }
        public Entity Consumer { get; set; }
        public string AMCPlanName { get; set; }
    }
}
