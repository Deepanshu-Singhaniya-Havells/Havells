using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.AMCInvoiceSync.Model
{
    public class SMSRequestModel
    {
        public string MobileNumber { get; set; }
        public string Message { get; set; }
    }
    public class SMSResponseModel
    {
        public string ResponseCode { get; set; }
        public string ResponseMessage { get; set; }
        public string Data { get; set; }
    }
}
