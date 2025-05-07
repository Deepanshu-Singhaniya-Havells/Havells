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
    public class AuthenticateCustomerOTP
    {
        [DataMember]
        public string UName { get; set; }
        [DataMember]
        public string OTP { get; set; }
        [DataMember]
        public string Method { get; set; }
        [DataMember(IsRequired = false)]
        public string ContGuid { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }
        [DataMember(IsRequired = false)]
        public bool IfValidated { get; set; }
        [DataMember(IsRequired = false)]
        public string FirstName { get; set; }
        [DataMember(IsRequired = false)]
        public string LastName { get; set; }
        [DataMember(IsRequired = false)]
        public bool NewAPK { get; set; }
        [DataMember(IsRequired = false)]
        public string MobileNumber { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCode { get; set; }
        [DataMember(IsRequired = false)]
        public int ERROR_CODE { get; set; }
        public AuthenticateCustomerOTP ValidateThisCustomerOTP(AuthenticateCustomerOTP bridge)
        {
            NewAPK = true;
            #region Not in use
            //IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            //try
            //{

            //    QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            //    Query.ColumnSet = new ColumnSet("emailaddress1", "hil_password", "firstname", "lastname", "mobilephone");
            //    Query.Criteria = new FilterExpression(LogicalOperator.And);
            //    Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, bridge.UName);
            //    EntityCollection Found = service.RetrieveMultiple(Query);
            //    if (Found.Entities.Count == 1)
            //    {
            //        Contact Cont = (Contact)Found.Entities[0];
            //        bridge.ContGuid = Cont.ContactId.Value.ToString();
            //        if (Cont.MobilePhone != null)
            //            bridge.MobileNumber = Cont.MobilePhone;
            //        bridge.IfValidated = true;
            //        if (Cont.LastName != null)
            //            bridge.LastName = Cont.LastName;
            //        if (Cont.FirstName != null)
            //            bridge.FirstName = Cont.FirstName;
            //        bridge.PinCode = GetPermanentAddressPinCode(service, Cont.Id);
            //        int iOTPLength = 6;
            //        string[] saAllowedCharacters = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            //        bridge.OTP = GenerateRandomOTP(iOTPLength, saAllowedCharacters);
            //        if (NewAPK)
            //            SendSMSToCustomer(bridge.OTP, bridge.FirstName, bridge.LastName, bridge.MobileNumber);
            //        else
            //        {
            //            //OUTGOING_SMS iOSms = new OUTGOING_SMS();
            //            //iOSms.MESSAGE = "Dear Customer, We have launched New Application. Please Uninstall Current App & Reinstall from below link << >>";
            //            //iOSms.TO = bridge.MobileNumber;
            //            //iOSms.OUTGOINGSMSMETHOD(iOSms);
            //        }
            //        //SendEmailToCustomer(bridge.OTP, bridge.FirstName, bridge.LastName, bridge.MobileNumber);
            //        bridge.status = true;
            //    }
            //    else
            //    {
            //        QueryExpression Query1 = new QueryExpression(Contact.EntityLogicalName);
            //        Query1.ColumnSet = new ColumnSet("emailaddress1", "hil_password", "firstname", "lastname", "mobilephone");
            //        Query1.Criteria = new FilterExpression(LogicalOperator.And);
            //        Query1.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, bridge.UName);
            //        EntityCollection Found1 = service.RetrieveMultiple(Query1);
            //        if (Found1.Entities.Count == 1)
            //        {
            //            Contact Cont = (Contact)Found1.Entities[0];
            //            bridge.ContGuid = Cont.ContactId.Value.ToString();
            //            if (Cont.MobilePhone != null)
            //                bridge.MobileNumber = Cont.MobilePhone;
            //            bridge.IfValidated = true;
            //            if (Cont.LastName != null)
            //                bridge.LastName = Cont.LastName;
            //            if (Cont.FirstName != null)
            //                bridge.FirstName = Cont.FirstName;
            //            bridge.PinCode = GetPermanentAddressPinCode(service, Cont.Id);
            //            int iOTPLength = 6;
            //            string[] saAllowedCharacters = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            //            bridge.OTP = GenerateRandomOTP(iOTPLength, saAllowedCharacters);
            //            if (NewAPK) SendSMSToCustomer(bridge.OTP, bridge.FirstName, bridge.LastName, bridge.MobileNumber);
            //            else
            //            {
            //                //OUTGOING_SMS iOSms = new OUTGOING_SMS();
            //                //iOSms.MESSAGE = "Dear Customer, We have launched New Application. Please Uninstall Current App & Reinstall from below link << >>";
            //                //iOSms.TO = bridge.MobileNumber;
            //                //iOSms.OUTGOINGSMSMETHOD(iOSms);
            //            }
            //            bridge.status = true;
            //            bridge.ERROR_CODE = 0;
            //        }
            //        else
            //        {
            //            bridge.ContGuid = "NOT FOUND";
            //            bridge.MobileNumber = "NOT FOUND";
            //            bridge.IfValidated = false;
            //            bridge.LastName = "NOT FOUND";
            //            bridge.FirstName = "NOT FOUND";
            //            bridge.PinCode = "NOT FOUND";
            //            bridge.OTP = "";
            //            bridge.status = false;
            //            bridge.ERROR_CODE = 1;
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    bridge.ContGuid = ex.Message.ToUpper();
            //    bridge.MobileNumber = "NOT FOUND";
            //    bridge.IfValidated = false;
            //    bridge.LastName = "NOT FOUND";
            //    bridge.FirstName = "NOT FOUND";
            //    bridge.PinCode = "NOT FOUND";
            //    bridge.OTP = "";
            //    bridge.status = false;
            //}
            #endregion

            bridge.ContGuid = "Access Denied!!!.";
            bridge.MobileNumber = "NOT FOUND";
            bridge.IfValidated = false;
            bridge.LastName = "NOT FOUND";
            bridge.FirstName = "NOT FOUND";
            bridge.PinCode = "NOT FOUND";
            bridge.OTP = "";
            bridge.status = false;
            return (bridge);
        }
        public string GetPermanentAddressPinCode(IOrganizationService service, Guid ContId)
        {
            string Pin = "";
            QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_pincode");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, ContId);
            Query.Criteria.AddCondition("hil_addresstype", ConditionOperator.Equal, 1);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                hil_address iAddress = Found.Entities[0].ToEntity<hil_address>();
                Pin = iAddress.hil_PinCode.Name;
            }
            return Pin;
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
        public static void SendSMSToCustomer(string OTP, string First_Name, string Last_Name, string Mobile)
        {
            OUTGOING_SMS iOSms = new OUTGOING_SMS();
            iOSms.MESSAGE = "Dear " + First_Name + " " + Last_Name + ", Please use " + OTP + " OTP for Havells Consumer App Login - Havells";
            iOSms.TO = Mobile;
            iOSms.SMSTEMPLATEID = "1107161709071483878";
            iOSms.OUTGOINGSMSMETHOD(iOSms);
        }
        //public static void SendEmailToCustomer(string OTP, string First_Name, string Last_Name, string Mobile)
        //{
        //    Email iEmail = new Email();
        //    //iEmail.
        //}
    }
}
