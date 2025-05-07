using HavellsSync_ModelData.EasyReward;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.Epos
{
    public class EposUserinfoDetails
    {
        public string ConsumerName { get; set; }
        public string Gender { get; set; }
        public string EmailId { get; set; }
        public string DOB { get; set; }       
        public string Consent { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string PINCode { get; set; }
        public bool Status { get; set; } = false;
        public string Response { get; set; }


    }

    public class EposJobStatus
    {
        public string JobStatus { get; set; }      
        public bool Status { get; set; } = false;
        public string Response { get; set; }

    }

    public class ServiceCallRequestData
    {
        public string ConsumerName { get; set; }
        public string EmailId { get; set; }
        public string MobileNumber { get; set; }
        public DateTime DOB { get; set; }
        public string Consent { get; set; }
        public string Gender { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string Landmark { get; set; }
        public string PINCode { get; set; }       
        public Guid GuidID { get; set; }
        public string Response { get; set; }
        public bool Status { get; set; }
        public List<ServiceCallLineItemData> ServiceCallLineItem { get; set; }
    }

    public class ServiceCallLineItemData
    {
        public string SerialNumber { get; set; }
        public string SKUCode { get; set; }
        public string PreferredDateofService { get; set; } 
        public string PreferredTimeofService { get; set; }
        public string InstallationRequired { get; set; }
    }

    public class EposLoyaltyparam<T>
    {
        public T data { get; set; }
    }

    public class ClsMobileNumber
    {
        public string MobileNumber { get; set; }
    }

    public class ClsJobId
    {
        public string JobId { get; set; }
    }

    public class ResponseMessage
    {
        public string Response { get; set; } 
    }
}
