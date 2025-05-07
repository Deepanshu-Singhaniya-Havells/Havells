using System;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    public class ReturnResult
    {
        [DataMember(IsRequired = false)]
        public string StatusCode { get; set; }

        [DataMember(IsRequired = false)]
        public string StatusDescription { get; set; }

        [DataMember(IsRequired = false)]
        public Guid CustomerGuid { get; set; }

        [DataMember(IsRequired = false)]
        public string MobileNumber { get; set; }

        [DataMember(IsRequired = false)]
        public string CustomerName { get; set; }

        [DataMember(IsRequired = false)]
        public string EmailId { get; set; }
        [DataMember]
        public bool? Consent { get; set; }
        [DataMember]
        public bool? SubscribeForMsgService { get; set; } // hil_subscribeformessagingservice
        [DataMember]
        public string PreferredLanguage { get; set; } // hil_preferredlanguageforcommunication

        [DataMember(IsRequired = false)]
        public string PINCode { get; set; }
        [DataMember(IsRequired = false)]
        public string Address { get; set; }
    }
}
