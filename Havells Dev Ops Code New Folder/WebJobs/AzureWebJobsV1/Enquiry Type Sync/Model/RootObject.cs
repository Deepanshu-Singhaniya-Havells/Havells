using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Enquiry_Type_Sync.Model
{
    public class RootObject
    {
        public object Result { get; set; }
        public List<EnquiryTypeData> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class EnquiryTypeData {
        public string EnquiryTypeCode { get; set; }
        public string EnquiryTypeDesc { get; set; }
        public bool IsActive { get; set; }
    }
}
