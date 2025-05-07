using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using Havells_Plugin;
using System.Net.Http;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AuthenticationServiceExternal
    {
        [DataMember]
        public string UID { get; set; }
        [DataMember]
        public string CODE { get; set; }
        [DataMember(IsRequired = false)]
        public string SOURCE { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string Message { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public AuthenticationServiceExternal Authenticate(AuthenticationServiceExternal oSFA)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            string _api = string.Empty;
            string sUserName = string.Empty;
            string sPassword = string.Empty;

            string configurtionparameter = string.Empty;
            try
            {
                if (SOURCE.ToLower() == "sfa")
                {
                    configurtionparameter = "SFA_Integration_Outgoing";
                    byte[] bytearray = Encoding.UTF8.GetBytes("{\"UID\":\"" + oSFA.UID + "\"}");
                }
                if (SOURCE.ToLower() == "dealerportal")
                {
                    configurtionparameter = "DP_Integration_Outgoing";
                }

                var randomstring = string.Empty;
                StringBuilder builder = new StringBuilder();

                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = (from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                               where _IConfig.hil_name == configurtionparameter
                               select new
                               {
                                   _IConfig.hil_Url,
                                   _IConfig.hil_Aid,
                                   _IConfig.hil_Pin,
                                   _IConfig.hil_Signature,
                                   _IConfig.hil_Username,
                                   _IConfig.hil_Password
                               }).Take(1);
                    foreach (var iobj in obj)
                    {
                        WebRequest request = WebRequest.Create(iobj.hil_Url);
                        if (SOURCE.ToLower() == "sfa")
                        {
                            byte[] bytearray = Encoding.UTF8.GetBytes("{\"UID\":\"" + oSFA.UID + "\"}");

                            request.Method = "POST";
                            request.ContentType = "application/json";
                            request.ContentLength = bytearray.Length;
                            Stream datastream = request.GetRequestStream();
                            datastream.Write(bytearray, 0, bytearray.Length);
                            datastream.Close();
                        }
                        else if (SOURCE.ToLower() == "dealerportal")
                        {
                            request = WebRequest.Create(iobj.hil_Url + "?username=" + oSFA.UID);
                            request.Method = "POST";

                        }
                        WebResponse response = request.GetResponse();
                        Stream dataStream = Stream.Null;
                        Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                        string IfOkay = ((HttpWebResponse)response).StatusDescription;
                        dataStream = response.GetResponseStream();
                        StreamReader reader = new StreamReader(dataStream);
                        string responseFromServer = reader.ReadToEnd();
                        oSFA.RESULT = "SUCCESS";
                        oSFA.RESULT_DESC = responseFromServer;
                    }
                }
            }
            catch (Exception ex)
            {
                oSFA.RESULT = "FAILURE";
                oSFA.RESULT_DESC = ex.Message.ToUpper();
            }
            return (oSFA);
        }
    }
}
