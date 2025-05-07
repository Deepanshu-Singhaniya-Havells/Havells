using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells_Plugin.HelperIntegration
{
    class EMSDataHelper
    {
    }

    [System.Runtime.Serialization.DataContractAttribute()]
    public class EMS_Request
    {
        [System.Runtime.Serialization.DataMemberAttribute()]
        public Data Data { get; set; }
    }

    [System.Runtime.Serialization.DataContractAttribute()]
    public class Data
    {
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string TicketNo { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string FirstName { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string LastName { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string Address { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string Email { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string MobileNo { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string PinCode { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string DivisionCode { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string EmployeeType { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string EnquiryType { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string Interest { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string TicketParentId { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string Remarks { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string UTMCode { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string CampaignCode { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string CampaignDesc { get; set; }
    }

    [System.Runtime.Serialization.DataContractAttribute()]
    public class EMS_Response
    {
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string ResponseStatusCode { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string Message { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string ServiceToken { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public List<Result> Results { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string IsDeleteInsert { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string EarlierChannelType { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string MaxLastSyncDate { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string IsSuccess { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string DataPacket { get; set; }
    }

    [System.Runtime.Serialization.DataContractAttribute()]
    public class Result
    {
        [System.Runtime.Serialization.DataMemberAttribute()]
        public string TicketNo { get; set; }

        [System.Runtime.Serialization.DataMemberAttribute()]
        public string ErrorMessage { get; set; }
    }
}