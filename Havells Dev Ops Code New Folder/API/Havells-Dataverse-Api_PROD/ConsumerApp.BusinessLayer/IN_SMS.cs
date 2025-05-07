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
    public class IN_SMS
    {
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public IN_SMS CreateInComingSMS(string FROM, string TO, string MESSAGE, int JobType)
        {
            IN_SMS oSMS = new IN_SMS();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                hil_smsconfiguration InSMS = new hil_smsconfiguration();
                InSMS.hil_Message = MESSAGE;
                InSMS.hil_MobileNumber = FROM;
                InSMS.hil_Direction = new OptionSetValue(1);
                InSMS["hil_tomobile"] = TO;
                if(JobType == 1 || JobType == 2)
                    InSMS["hil_jobtype"] = new OptionSetValue(JobType);
                service.Create(InSMS);
                oSMS.RESULT = "SUCCESS";
                oSMS.RESULT_DESC = "SUCCESS";
            }
            catch(Exception ex)
            {
                oSMS.RESULT = "FAILURE";
                oSMS.RESULT_DESC = ex.Message.ToUpper();
            }
            return (oSMS);
        }
    }
}
