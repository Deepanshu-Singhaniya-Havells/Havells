using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class LoyaltyProgram
    {
        public sendEnrollmentRes SendEnrollmentLink(sendEnrollmentReq sendEnrollmentReq)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");

                if (string.IsNullOrWhiteSpace(sendEnrollmentReq.mobileNo))
                {
                    return new sendEnrollmentRes { status = false, message = "Mobile No. is required." };
                }
                else if (!Regex_MobileNo.IsMatch(sendEnrollmentReq.mobileNo))
                {
                    return new sendEnrollmentRes { status = false, message = "Invalid Mobile No." };
                }
                if (string.IsNullOrWhiteSpace(sendEnrollmentReq.jobGuid))
                {
                    return new sendEnrollmentRes { status = false, message = "Job Guid is required." };
                }
                else if (new Guid(sendEnrollmentReq.jobGuid) == Guid.Empty)
                {
                    return new sendEnrollmentRes { status = false, message = "Job Guid is required" };
                }
                if (service != null)
                {
                    EntityReference entTemplate = new EntityReference("hil_smstemplates", new Guid("96db1c6d-674d-ef11-acce-6045bde7b2f4"));
                    EntityReference entJob = new EntityReference("msdyn_workorder", new Guid(sendEnrollmentReq.jobGuid));
                    string SMSSubject = string.Empty;
                    string SMSBody = string.Empty;
                    if (entTemplate.Id != Guid.Empty)
                    {
                        Entity entSMSTemplate = service.Retrieve("hil_smstemplates", entTemplate.Id, new ColumnSet("hil_name", "hil_templatebody"));
                        SMSSubject = entSMSTemplate.GetAttributeValue<string>("hil_name");
                        SMSBody = entSMSTemplate.GetAttributeValue<string>("hil_templatebody");

                    }
                    string url = "https://bit.ly/3Wy9uoq";
                    string discountPercent = "3%";
                    Entity smsEntity = new Entity("hil_smsconfiguration");
                    smsEntity["hil_smstemplate"] = entTemplate;
                    smsEntity["subject"] = SMSSubject;
                    smsEntity["hil_message"] = "Congrats, you have earned Havells Happiness loyalty membership! SignUp now " + url + " to unlock upto " + discountPercent + " points & priority service. T&C - Havells";
                    smsEntity["hil_mobilenumber"] = sendEnrollmentReq.mobileNo;
                    smsEntity["hil_direction"] = new OptionSetValue(2);
                    smsEntity["regardingobjectid"] = entJob;
                    Guid id = service.Create(smsEntity);
                    return new sendEnrollmentRes { Id = id.ToString(), status = true, message = "Success." };
                }
                else
                {
                    return new sendEnrollmentRes { status = false, message = "D365 service unavailable..." };
                }
            }
            catch (Exception ex)
            {
                return new sendEnrollmentRes { status = false, message = ex.Message };
            }
        }
    }
    public class sendEnrollmentReq
    {
        [DataMember]
        public string mobileNo { get; set; }
        [DataMember]
        public string jobGuid { get; set; }
    }
    public class sendEnrollmentRes
    {
        [DataMember]
        public string Id { get; set; }

        [DataMember]
        public string message { get; set; }
        [DataMember]
        public bool status { get; set; }
    }

}
