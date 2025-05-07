//CHANNEL - 1 - EMAIL
//          2 - SMS

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
    public class ForgotPasswordCheck1
    {
        [DataMember]
        public int CHANNEL { get; set; }
        [DataMember]
        public string CONT_GUID { get; set; }
        [DataMember(IsRequired = false)]
        public string OTP { get; set; }
        [DataMember(IsRequired = false)]
        public bool NewAPK { get; set; }
        public ForgotPasswordCheck1 SendOTPFromService(ForgotPasswordCheck1 bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            Contact iContact = (Contact)service.Retrieve(Contact.EntityLogicalName, new Guid(bridge.CONT_GUID),
                   new ColumnSet(false));
            NewAPK = true;
            if (!NewAPK) throw new Exception();
            try
            {               
                int iOTPLength = 6;
                string[] saAllowedCharacters = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
                string iOTP = GenerateRandomOTP1(iOTPLength, saAllowedCharacters);
                bridge.OTP = iOTP;
                iContact.hil_OTP = iOTP;
                iContact["hil_channel"] = Convert.ToInt32(bridge.CHANNEL);
                service.Update(iContact);
            }
            catch (Exception ex)
            {
                bridge.OTP = ex.Message.ToUpper();
                if (NewAPK)
                {
                    iContact.hil_OTP = "Dear Customer, We have launched New Application. Please Uninstall Current App & Reinstall from below link << >>"; ;
                    iContact["hil_channel"] = Convert.ToInt32(bridge.CHANNEL);
                    service.Update(iContact);
                }
            }
            return (bridge);
        }
        public static string GenerateRandomOTP1(int iOTPLength, string[] saAllowedCharacters)
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
    }
}
