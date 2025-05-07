using Microsoft.Exchange.WebServices.Data;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;

namespace SMTP_Integration
{
    class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();

            var SMTPUser = ConfigurationManager.AppSettings["SMTPUser"].ToString();
            var SMTPUserPassword = ConfigurationManager.AppSettings["SMTPUserPassword"].ToString();
            var SMTPURL = ConfigurationManager.AppSettings["SMTPURL"].ToString();
            var subject = ConfigurationManager.AppSettings["subject"].ToString();
            var ApproverMailID = ConfigurationManager.AppSettings["ApproverMailID"].ToString();

            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService _D365Service = HavellsConnection.CreateConnection.createConnection(finalString);

            retriveMail(_D365Service, SMTPUser, SMTPUserPassword, SMTPURL, subject, ApproverMailID);
            Console.WriteLine("Email done.");
            
            //Random rnd = new Random();
            //int num = rnd.Next(100, 1000);
            //for(int i =0; i < num; i++)
            //{
            //    Console.WriteLine(DateTime.Now.ToString());
            //}
        }
        public static void retriveMail(IOrganizationService _D365Service, string SMTPUser, string SMTPUserPassword, string SMTPURL, string subjectCont, string ApproverMailID)
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            
            service.Credentials = new WebCredentials(SMTPUser, SMTPUserPassword);
            service.Url = new Uri(SMTPURL);
            service.PreAuthenticate = true;
            
            ItemView view = new ItemView(int.MaxValue);
            int count = 0;
            foreach (EmailMessage email in service.FindItems(WellKnownFolderName.Inbox, SetFilter(), view))
            {
                if (email.Subject.Contains(subjectCont))
                {
                    if (!email.IsRead)
                    {
                        count++;
                        Console.WriteLine(count);
                        email.Load(new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.TextBody));
                        string recipients = "";

                        foreach (EmailAddress emailAddress in email.CcRecipients)
                        {
                            recipients += ";" + emailAddress.Address.ToString();
                        }
                        string internetMessageId = email.InternetMessageId;
                        string fromAddress = email.From.Address;
                        string recipient = recipients;
                        string subject = email.Subject;
                        string mailbody = email.Body;

                        string[] subArr = subject.Split('|');
                        string OANumber = subArr[0].Replace("OA ", "");
                        if (email.From.Address.ToLower() == ApproverMailID.ToLower())
                        {
                            QueryExpression query = new QueryExpression("hil_oaheader");


                            query.ColumnSet = new ColumnSet("hil_approvalstatus");
                            query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, OANumber.Trim()));
                            EntityCollection _OAColl = _D365Service.RetrieveMultiple(query);
                            if (_OAColl.Entities.Count == 1)
                            {
                                if (_OAColl[0].GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value == 3)
                                {
                                    int approvalStatus = subject.Contains("Approved") ? 1 : (subject.Contains("Rejected") ? 2 : 0);
                                    if (approvalStatus != 0)
                                    {
                                        System.Text.RegularExpressions.Regex rx = new System.Text.RegularExpressions.Regex("<[^>]*>");
                                        string str = rx.Replace(email.Body, "");
                                        str = str.Replace("\r\n\r\n\r\n\r\n\r\n\r\nRemarks: ", "");
                                        Entity oaHeader = new Entity(_OAColl.EntityName, _OAColl[0].Id);
                                        oaHeader["hil_approvalstatus"] = new OptionSetValue(approvalStatus);
                                        oaHeader["hil_approverremarks"] = str;
                                        _D365Service.Update(oaHeader);
                                        email.IsRead = true;
                                        email.Update(ConflictResolutionMode.AutoResolve);
                                    }
                                }
                            }
                        }
                    }
                }
                email.IsRead = true;
                email.Update(ConflictResolutionMode.AutoResolve);
            }
        }
        private static SearchFilter SetFilter()
        {
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            //searchFilterCollection.Add(new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, "APPROVAL: LIMIT ON SELF EXPOSURE. Customer Name-"));
            // searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true));
            SearchFilter s = new SearchFilter.SearchFilterCollection(Microsoft.Exchange.WebServices.Data.LogicalOperator.Or, searchFilterCollection.ToArray());
            return s;
        }

        private static bool CertificateValidationCallBack(
         object sender,
         System.Security.Cryptography.X509Certificates.X509Certificate certificate,
         System.Security.Cryptography.X509Certificates.X509Chain chain,
         System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }
    }
}
