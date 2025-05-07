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
    public class OUTGOING_SMS
    {
        [DataMember]
        public string MESSAGE { get; set; }
        [DataMember]
        public string SMSTEMPLATEID { get; set; }
        [DataMember]
        public string TO { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT { get; set; }
        [DataMember(IsRequired = false)]
        public string RESULT_DESC { get; set; }
        public OUTGOING_SMS OUTGOINGSMSMETHOD(OUTGOING_SMS oSMS)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            string _api = string.Empty;
            string sUserName = string.Empty;
            string sPassword = string.Empty;


            try
            {

                var randomstring = string.Empty;
                StringBuilder builder = new StringBuilder();
                Random random = new Random();
                int ch;
                for (int i = 0; i < 6; i++)
                {
                    ch = random.Next(9);
                    builder.Append(ch);
                }
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = (from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                               where _IConfig.hil_name == "SMS_Integration_Outgoing"
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
                        //if (iobj.hil_Url != null)
                        //    _api = iobj.hil_Url;
                        //if (iobj.hil_Aid != null)
                        //    _api = _api + iobj.hil_Aid;
                        //if (iobj.hil_Pin != null)
                        //    _api = _api + "&pin=" + iobj.hil_Pin + "&mnumber=";
                        //if (oSMS.TO != null)
                        //    _api = _api + oSMS.TO + "&message=";
                        //if (oSMS.MESSAGE != null)
                        //    _api = _api + oSMS.MESSAGE + "&signature=";
                        //else
                        //    _api = _api + builder.ToString() + "&signature=";
                        //if (iobj.hil_Signature != null)
                        //    _api = _api + iobj.hil_Signature;

                        if (iobj.hil_Username != null)
                            sUserName = iobj.hil_Username;
                        if (iobj.hil_Password != null)
                            sPassword = iobj.hil_Password;

                        _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=" + oSMS.TO + "&signature=HAVELL&message=" + oSMS.MESSAGE + "&dlt_entity_id=110100001483&dlt_template_id=" + oSMS.SMSTEMPLATEID;

                        WebRequest request = WebRequest.Create(_api);
                        string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                        request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                        request.Method = "GET";
                        WebResponse response = request.GetResponse();
                        Stream dataStream = Stream.Null;
                        Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                        string IfOkay = ((HttpWebResponse)response).StatusDescription;
                        dataStream = response.GetResponseStream();
                        StreamReader reader = new StreamReader(dataStream);
                        string responseFromServer = reader.ReadToEnd();
                        oSMS.RESULT = IfOkay;
                        oSMS.RESULT_DESC = responseFromServer;
                        oSMS.MESSAGE = builder.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                oSMS.RESULT = "FAILURE";
                oSMS.RESULT_DESC = ex.Message.ToUpper();
            }
            return (oSMS);
        }
    }
}
