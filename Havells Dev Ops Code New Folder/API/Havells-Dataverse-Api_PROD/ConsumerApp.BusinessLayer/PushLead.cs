using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Deployment;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class PUSH_LEAD
    {
        [DataMember(IsRequired = false)]
        public string MOBILE { get; set; }
        [DataMember(IsRequired = false)]
        public string DIVISION { get; set; }
        [DataMember(IsRequired = false)]
        public string PIN { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public PUSH_LEAD CreateLeadForPreSalesEnquiry(PUSH_LEAD Lead)
        {
            PUSH_LEAD oLead = new PUSH_LEAD();
            oLead.DIVISION = "";
            oLead.MOBILE = "";
            oLead.PIN = "";
            oLead.RESULT = "SUCCESS";
            oLead.RESULT_DESC = "SUCCESS"; 
            return (oLead);
        }
    }
}
