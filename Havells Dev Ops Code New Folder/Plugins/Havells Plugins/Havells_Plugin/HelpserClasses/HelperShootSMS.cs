using System;
using System.IO;
using System.Net;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Security.Cryptography;
using System.Collections.Generic;
using Newtonsoft.Json;
// Microsoft Dynamics CRM namespace(s)

namespace Havells_Plugin
{
    public class HelperShootSMS
    {
        public static string EncriptionKey = "12s2s121sasfdasdf45346fwrt3w56fw";
        public static string _smsToTechnician = "1107162184423049071", _smsToConsumer = "1107162314614376909";

        public static void OnDemandSMSShootFunction(IOrganizationService service, string Msg, string Mob, hil_smsconfiguration Conf)
        {
            try
            {
                string _api = string.Empty;
                string sUserName = string.Empty;
                string sPassword = string.Empty;
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = (from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                              where _IConfig.hil_name == "SMS_Integration_Outgoing"
                              select new {
                                  _IConfig.hil_Url,
                                  _IConfig.hil_URL2,
                                  _IConfig.hil_Aid,
                                  _IConfig.hil_Pin,
                                  _IConfig.hil_Signature,
                                  _IConfig.hil_Username,
                                  _IConfig.hil_Password,
                                  _IConfig.hil_OrgName}).Take(1);
                    foreach(var iobj in obj)
                    {
                        #region Existing API
                        if (iobj.hil_Url != null)
                            _api = iobj.hil_Url;
                        if (iobj.hil_Aid != null)
                            _api = _api + iobj.hil_Aid;
                        if (iobj.hil_Pin != null)
                            _api = _api + "&pin=" + iobj.hil_Pin + "&mnumber=";
                        if (Mob != null)
                            _api = _api + Mob + "&message=";
                        if (Msg != null)
                            _api = _api + Msg + "&signature=";
                        if (iobj.hil_Signature != null)
                            _api = _api + iobj.hil_Signature;
                        if (iobj.hil_Username != null)
                            sUserName = iobj.hil_Username;
                        if (iobj.hil_Password != null)
                            sPassword = iobj.hil_Password;
                        //WebResponse response;
                        WebRequest request = WebRequest.Create(_api);
                        string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                        request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                        request.Method = "GET";
                        WebResponse response = null;
                        string IfOkay = string.Empty;
                        string responseFromServer = string.Empty;
                        try
                        {
                            response = request.GetResponse();
                            Stream dataStream = Stream.Null;
                            //Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                            IfOkay = ((HttpWebResponse)response).StatusDescription;
                            dataStream = response.GetResponseStream();
                            StreamReader reader = new StreamReader(dataStream);
                            responseFromServer = reader.ReadToEnd();
                        }
                        catch (WebException ex)
                        {
                            Conf.StatusCode = new OptionSetValue(910590001);
                            Conf["hil_responsefromserver"] = "GATEWAY BUSY";
                        }
                        finally
                        {

                            Conf["hil_responsefromserver"] = responseFromServer;
                            if (IfOkay == "OK")
                                Conf.StatusCode = new OptionSetValue(910590000); //Sent
                            else
                                Conf.StatusCode = new OptionSetValue(910590001); //Not Sent
                        }
                        #endregion
                        #region New API
                        //WebResponse response = null;
                        //string responseFromServer = string.Empty;
                        //if (iobj.hil_URL2 != null)
                        //    _api = iobj.hil_URL2;
                        //if (Mob != null)
                        //    _api = _api + Mob + "&send=alerts&text=";
                        //if (Msg != null)
                        //    _api = _api + Msg;
                        //WebRequest request1 = WebRequest.Create(_api);
                        ////string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                        ////request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                        //request1.Method = "GET";
                        //WebResponse response1 = null;
                        //string IfOkay1 = string.Empty;
                        //string responseFromServer1 = string.Empty;
                        //try
                        //{
                        //    response1 = request1.GetResponse();
                        //    Stream dataStream = Stream.Null;
                        //    //Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                        //    IfOkay1 = ((HttpWebResponse)response).StatusDescription;
                        //    dataStream = response1.GetResponseStream();
                        //    StreamReader reader = new StreamReader(dataStream);
                        //    responseFromServer = reader.ReadToEnd();
                        //}
                        //catch (WebException ex)
                        //{
                        //    Conf.StatusCode = new OptionSetValue(910590001);
                        //    Conf["hil_responsefromserver"] = "GATEWAY BUSY";
                        //}
                        //finally
                        //{

                        //    Conf["hil_responsefromserver"] = responseFromServer1;
                        //    if (IfOkay1 == "OK")
                        //        Conf.StatusCode = new OptionSetValue(910590000); //Sent
                        //    else
                        //        Conf.StatusCode = new OptionSetValue(910590001); //Not Sent
                        //}
                        #endregion
                        /* public string Get(string uri)
                           {
                               HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
                               request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                               using(HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                               using(Stream stream = response.GetResponseStream())
                               using(StreamReader reader = new StreamReader(stream))
                               {
                                    return reader.ReadToEnd();
                               }
                          }
                       */
                    }
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelperShootSMS : "+ex.Message.ToUpper());
            }
        }

        public static hil_smsconfiguration OnDemandSMSShootFunctionDLT(IOrganizationService service, string Msg, string Mob, hil_smsconfiguration Conf,string _templateId, string _custref, Guid _RegardingObjectId)
        {
            try
            {
                string _api = string.Empty;
                string sUserName = string.Empty;
                string sPassword = string.Empty;
                
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = (from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                               where _IConfig.hil_name == "SMS_Integration_Outgoing"
                               select new
                               {
                                   _IConfig.hil_Url,
                                   _IConfig.hil_URL2,
                                   _IConfig.hil_Aid,
                                   _IConfig.hil_Pin,
                                   _IConfig.hil_Signature,
                                   _IConfig.hil_Username,
                                   _IConfig.hil_Password,
                                   _IConfig.hil_OrgName
                               }).Take(1);
                    foreach (var iobj in obj)
                    {
                        if (iobj.hil_Username != null)
                            sUserName = iobj.hil_Username;
                        if (iobj.hil_Password != null)
                            sPassword = iobj.hil_Password;

                        if (_templateId == "1107161191448698079" || _templateId == "1107161191438154934") //KKG Code Send and Resend Templates
                        {

                            //return;
                            //_api = "https://digimate.airtel.in:15443/BULK_API/SendMessage?loginID=havells_htu2&password=havells@123&mobile=" + Mob + "&text=" + Msg + "&senderid=HAVELL&DLT_TM_ID=1001096933494158&DLT_CT_ID=" + _templateId + "&DLT_PE_ID=110100001483&route_id=DLT_SERVICE_IMPLICT&Unicode=0&camp_name=havells_u";
                            AirtelIQRequest airtelIQRequest = new AirtelIQRequest()
                            {
                                customerId = "Havells",
                                destinationAddress = new List<string>() { Mob },
                                message = Msg,
                                sourceAddress = "HAVELL",
                                messageType = "SERVICE_IMPLICIT",
                                dltTemplateId = _templateId,
                                metaData = new MetaData { Key = "Value" }
                            };

                            string BaseUrl = "https://openapi.airtel.in/gateway/airtel-iq-sms-utility/sendSms";
                            airtelIQRequest.entityId = "110100001483";
                            var json = JsonConvert.SerializeObject(airtelIQRequest);

                            try
                            {
                                WebRequest webRequest = WebRequest.Create(BaseUrl);
                                webRequest.Method = "POST";
                                byte[] data = Encoding.ASCII.GetBytes(json);
                                webRequest.ContentType = "application/json";
                                webRequest.ContentLength = data.Length;
                                webRequest.Headers.Add("Authorization", "Basic UG9saWN5QmF6YWFyOkZtQThGSFYzQEJ3RHUmNms=");
                                Stream requestStream = webRequest.GetRequestStream();
                                requestStream.Write(data, 0, data.Length);
                                WebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
                                string IfOkay = ((HttpWebResponse)webResponse).StatusCode.ToString();
                                Stream responseStream = webResponse.GetResponseStream();
                                StreamReader myStreamReader = new StreamReader(responseStream, Encoding.Default);
                                string responseFromServer = myStreamReader.ReadToEnd();

                                if (IfOkay == "OK")
                                {
                                    Conf["statuscode"] = new OptionSetValue(910590000); //Sent
                                    Conf["hil_responsefromserver"] = "SUCCESS";
                                }
                                else
                                {
                                    Conf["statuscode"] = new OptionSetValue(910590001); //Not Sent
                                    Conf["hil_responsefromserver"] = responseFromServer;
                                }
                                myStreamReader.Close();
                                responseStream.Close();
                                webResponse.Close();
                            }
                            catch (WebException we)
                            {
                                Conf.StatusCode = new OptionSetValue(910590001);
                                Conf["hil_responsefromserver"] = "GATEWAY BUSY API Url " + BaseUrl + " PayLoad: " + json;
                            }
                        }
                        else
                        {
                            if ((_templateId == _smsToConsumer || _templateId == _smsToTechnician) && _RegardingObjectId != Guid.Empty)
                            {
                                Entity _entJob = service.Retrieve("msdyn_workorder", _RegardingObjectId, new ColumnSet("ownerid"));
                                if (DoesUserHasRole(_entJob.GetAttributeValue<EntityReference>("ownerid").Id, service))
                                {
                                    Msg = ChangeMobileNumber(Msg, _templateId);
                                }
                            }
                            
                            _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=" + Mob + "&signature=HAVELL&message=" + Msg + "&dlt_entity_id=110100001483&dlt_template_id=" + _templateId + "&cust_ref=" + _custref;

                            WebRequest request = WebRequest.Create(_api);
                            if (_templateId != "1107161191448698079" && _templateId != "1107161191438154934") //KKG Code Send and Resend Templates
                            {
                                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                                request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                            }
                            request.Method = "POST";
                            WebResponse response = null;
                            string IfOkay = string.Empty;
                            string responseFromServer = string.Empty;
                            try
                            {
                                response = request.GetResponse();
                                Stream dataStream = Stream.Null;
                                IfOkay = ((HttpWebResponse)response).StatusDescription;
                                dataStream = response.GetResponseStream();
                                StreamReader reader = new StreamReader(dataStream);
                                responseFromServer = reader.ReadToEnd();
                            }
                            catch (WebException ex)
                            {
                                Conf.StatusCode = new OptionSetValue(910590001);
                                Conf["hil_responsefromserver"] = "GATEWAY BUSY APIPayLoad: " + _api;
                                Conf["hil_message"] = Msg;
                            }
                            finally
                            {

                                Conf["hil_responsefromserver"] = "Plugin: " + responseFromServer;
                                Conf["hil_message"] = Msg;
                                if (IfOkay == "OK")
                                    Conf.StatusCode = new OptionSetValue(910590000); //Sent
                                else
                                    Conf.StatusCode = new OptionSetValue(910590001); //Not Sent
                            }


                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelperShootSMS : " + ex.Message.ToUpper());
            }
            return Conf;
        }

        public static string ChangeMobileNumber(string message, string template)
        {
            string toMatch = "";
            char endingMatch = '~';
            if (template == _smsToConsumer)
            {
                toMatch = "mob no";
                endingMatch = '.';
            }
            else if (template == _smsToTechnician)
            {
                toMatch = "Mob No";
                endingMatch = 'A';
            }

            int start = 0;
            int j = 0;

            while (start < message.Length && j < toMatch.Length)
            {
                if (j == toMatch.Length)
                {
                    Console.WriteLine("hey I found it");
                    break;
                }
                else if (message[start] == toMatch[j])
                {
                    start++;
                    j++;
                }
                else
                {
                    start++;
                    j = 0;
                }
            }

            int end = start;
            start++;
            while (message[end] != endingMatch) end++;
            if (template == _smsToTechnician) end--;
            string updatedMessage = message.Remove(start, end - start);
            updatedMessage = updatedMessage.Insert(start, "08048251313");
            return updatedMessage;
        }
        public static bool DoesUserHasRole(Guid OwnerID, IOrganizationService service)
        {
            bool flag = false;
            string _roleName = "Franchise Call Masking";
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("name", "roleid"),
                LinkEntities = {
                    new LinkEntity
                    {
                        LinkFromEntityName = "role",
                        LinkFromAttributeName = "roleid",
                        LinkToEntityName = "systemuserroles",
                        LinkToAttributeName = "roleid",
                        LinkEntities =
                        {
                            new LinkEntity
                            {
                                LinkFromEntityName = "systemuserroles",
                                LinkFromAttributeName = "systemuserid",
                                LinkToEntityName = "systemuser",
                                LinkToAttributeName = "systemuserid",
                                EntityAlias = "ah",
                                LinkCriteria = new FilterExpression(LogicalOperator.And)
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression("systemuserid", ConditionOperator.Equal, OwnerID)
                                    }
                                }
                            }
                        }
                    }
                }
            };

            EntityCollection rolesColl = service.RetrieveMultiple(query);

            rolesColl.Entities.ToList().ForEach(e =>
            {
                string role = e.GetAttributeValue<string>("name");
                if (role == _roleName)
                {
                    flag = true;
                    return;
                }
            });
            return flag;
        }
        public static string EncryptString(string key, string plainText)
        {
            byte[] iv = new byte[16];
            byte[] array;

            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(key);
                aes.IV = iv;
                ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter streamWriter = new StreamWriter((Stream)cryptoStream))
                        {
                            streamWriter.Write(plainText);
                        }
                        array = memoryStream.ToArray();
                    }
                }
            }
            return Convert.ToBase64String(array);
        }

    }

    public class AirtelIQRequest
    {
        public string customerId { get; set; }
        public List<string> destinationAddress { get; set; }
        public string message { get; set; }
        public string sourceAddress { get; set; }
        public string messageType { get; set; }
        public string dltTemplateId { get; set; }
        public MetaData metaData { get; set; }
        public string entityId { get; set; }
    }

    public class MetaData
    {
        public string Key { get; set; }
        public string subAccountId { get; set; }
        public object createdBy { get; set; }
        public RbacSubAccount rbacSubAccount { get; set; }
        public string mdrCategory { get; set; }
    }

    public class RbacSubAccount
    {
        public string accountId { get; set; }
        public string status { get; set; }
        public Services services { get; set; }
    }
    public class Services
    {
        public AirtelIQSMS SMS { get; set; }
    }

    public class AirtelIQSMS
    {
        public bool creditFlag { get; set; }
        public string serviceStatus { get; set; }
        public int creditsAllotted { get; set; }
    }
}
