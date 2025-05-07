using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.AMCInvoiceSync.Model
{
    internal class PaymentAPIModel
    {
    }
    public class StatusRequest
    {
        
        public string PROJECT { get; set; }
        
        public string command { get; set; }
        
        public string var1 { get; set; }

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
    public class StatusResponse
    {
        
        public int status { get; set; }
        
        public string msg { get; set; }
        
        public List<TransactionDetail> transaction_details { get; set; }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
}
