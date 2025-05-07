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
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class RegisterCustomer
    {
        [DataMember]
        public string EMAIL { get; set; }
        [DataMember]
        public string PASSWORD { get; set; }
        [DataMember(IsRequired = false)]
        public int SALUTATION { get; set; }
        [DataMember(IsRequired = false)]
        public string FIRST_NAME { get; set; }
        [DataMember]
        public string LAST_NAME { get; set; }
        [DataMember]
        public string MOBILE { get; set; }
        [DataMember(IsRequired = false)]
        public string OTP { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string REC_ID { get; set; }//Output
        [DataMember(IsRequired = false)]
        public bool STATUS { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string STATUS_DESC {get; set;}//Output
        
        public RegisterCustomer InitiateCustomerCreation(RegisterCustomer bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            hil_consumerappbridge iBrd = new hil_consumerappbridge();
            if (bridge.MOBILE != null && bridge.PASSWORD != null && bridge.LAST_NAME != null && bridge.LAST_NAME != null &&
                bridge.EMAIL != null)
            {
                bool iExist = CheckIfContactExists(service, bridge.EMAIL, bridge.MOBILE);
                if(iExist == false)
                {
                    if (bridge.FIRST_NAME != null)
                        iBrd.hil_FirstName = bridge.FIRST_NAME;
                    iBrd.hil_LastName = bridge.LAST_NAME;
                    iBrd.hil_EmailId = bridge.EMAIL;
                    iBrd.hil_MobileNumber = bridge.MOBILE;
                    iBrd.hil_Password = bridge.PASSWORD;
                    iBrd["hil_salutationcode"] = (int)bridge.SALUTATION;
                    int iOTPLength = 6;
                    string[] saAllowedCharacters = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
                    bridge.OTP = GenerateRandomOTP(iOTPLength, saAllowedCharacters);
                    iBrd.hil_OTP = bridge.OTP;
                    Guid Rec = service.Create(iBrd);
                    bridge.REC_ID = Rec.ToString();
                    bridge.STATUS = true;
                    bridge.STATUS_DESC = "SUCCESS";
                    iBrd.hil_Description = "SUCCESS";
                }
                else
                {
                    bridge.OTP = "";
                    bridge.REC_ID = "";
                    bridge.STATUS = false;
                    bridge.STATUS_DESC = "ALREADY REGISTERED EMAIL/MOBILE";
                }
            }
            else
            {
                bridge.OTP = "";
                bridge.REC_ID = "";
                bridge.STATUS = false;
                bridge.STATUS_DESC = "MANDATORY DETAILS MISSED";
            }
            return (bridge);
        }
        public static string GenerateRandomOTP(int iOTPLength, string[] saAllowedCharacters)
        {
            string sOTP = String.Empty;
            string sTempChars = String.Empty;
            Random rand = new Random();
            for (int i = 0; i < iOTPLength; i++)
            {
                int p = rand.Next(0, saAllowedCharacters.Length);
                sTempChars = saAllowedCharacters[rand.Next(0, saAllowedCharacters.Length)];
                sOTP += sTempChars;
            }
            return sOTP;
        }
        public static bool CheckIfContactExists(IOrganizationService service, string Email, string Mobile)
        {
            bool iDuplicate = false;
            QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.Or);
            Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, Email);
            Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, Mobile);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDuplicate = true;
            }
            return iDuplicate;
        }
    }
}