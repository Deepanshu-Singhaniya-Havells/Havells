using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.GetChannelPartner
{
    #region GET PARTNER OBJECTS
    public class PartnerResult
    {
        public string sap_city_desc { get; set; }
        public string sap_dist_desc { get; set; }
        public string sap_state_desc { get; set; }
        public string dm_area_desc { get; set; }
        public string KUNNR { get; set; }
        public string KTOKD { get; set; }
        public string PAR_CODE { get; set; }
        public string KTONR { get; set; }
        public string VTXTM { get; set; }
        public string STREET { get; set; }
        public string STR_SUPPL3 { get; set; }
        public string ADDRESS3 { get; set; }
        public string dm_state { get; set; }
        public string dm_dist { get; set; }
        public string dm_city { get; set; }
        public string dm_pin { get; set; }
        public string KUKLA { get; set; }
        public string KNRZE { get; set; }
        public string NAME1 { get; set; }
        public string J_1IPANNO { get; set; }
        public string J_1ILSTNO { get; set; }
        public string BANKA { get; set; }
        public string BANKN { get; set; }
        public string BANKL { get; set; }
        public string SMTP_ADDR { get; set; }
        public string MOB_NUMBER { get; set; }
        public string PAR_NAME { get; set; }
        public string CONT_CODE { get; set; }
        public string CONT_PERSON { get; set; }
        public string STATUS { get; set; }
        public string DM_AREA { get; set; }
        public string SALES_STATUS { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
        public string RPMKR { get; set; }
        public string SAMPARK_MOB_NO { get; set; }
        public string FIRM_TYPE { get; set; }
        public string DM_REGION { get; set; }
        public string DM_BRANCH { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string AADHAAR_NO { get; set; }
        public string GST_NO { get; set; }
        public string MNC_NAME { get; set; }
        public string MNC_MOB_NO { get; set; }
        public string MNC_MAIL_ID { get; set; }
        public string KAM_ID { get; set; }
        public string BUSAB { get; set; }
        public string Longitude { get; set; }
        public string Latitude { get; set; }

    }
    public class PartnerRootObject
    {
        public object Result { get; set; }
        public List<PartnerResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class PartnerDivisionResult
    {
        public string KUNNR { get; set; }
        public string dm_ch { get; set; }
        public string dm_div { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string ERDAT { get; set; }
        public string STATUS { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string MODIFYBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MTIMESTAMP { get; set; }
        public string SAP_CH { get; set; }
        public string SAP_DIV { get; set; }
    }
    public class PartnerDivisionRootObject
    {
        public object Result { get; set; }
        public List<PartnerDivisionResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion
    public class GetPartner
    {
        #region GET PARTNER
        public static void GetPartnerAPIBackup(IOrganizationService service, string _syncDatetime)
        {
            try
            {
                string iKUNNR = string.Empty;
                Account iAccount = new Account();
                WebClient webClient = new WebClient();

                Integration intConf = GetIntegration(service, "Tier 1 Customer");
                //_enquiryDatetime = getTimeStamp(service);

                string uri = intConf.uri;
                string authInfo = intConf.userName + ":" + intConf.passWord;
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                webClient.Headers["Authorization"] = "Basic " + authInfo;

                Console.WriteLine("URL: " + uri);

                Console.WriteLine("EnquiryDatetime(YYYYMMDDHHMMSS): " + _syncDatetime);
                if (_syncDatetime != string.Empty && _syncDatetime.Trim().Length > 0)
                {
                    uri = uri + _syncDatetime;
                }
                else
                {
                    Console.WriteLine("Sync Datetime is mandetory.");
                    return;
                }

                Console.WriteLine("Downloading Channel Partner Data: " + DateTime.Now.ToString());
                var jsonData = webClient.DownloadData(uri);

                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(PartnerRootObject));
                PartnerRootObject rootObject = (PartnerRootObject)ser.ReadObject(new MemoryStream(jsonData));

                Console.WriteLine("Downloading Completed of Channel Partner Data: " + DateTime.Now.ToString());
                SyncChannelPartner(rootObject, service);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }

        public static void GetChannelPartnerData(IOrganizationService service, string _syncDatetime)
        {
            try
            {
                Integration intConf = GetIntegration(service, "Tier 1 Customer");
                string uri = intConf.uri;
                string authInfo = intConf.userName + ":" + intConf.passWord;
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));

                if (_syncDatetime != string.Empty && _syncDatetime.Trim().Length > 0)
                {
                    uri = uri + _syncDatetime;
                }
                else
                {
                    uri = uri + getTimeStamp(service);
                }
                Console.WriteLine("URL: " + uri);

                var client = new RestClient(uri);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Basic " + authInfo);
                Console.WriteLine("Downloading Channel Partner Data: " + DateTime.Now.ToString());
                IRestResponse response = client.Execute(request);
                PartnerRootObject rootObject = JsonConvert.DeserializeObject<PartnerRootObject>(response.Content);
                Console.WriteLine("Downloading Completed of Channel Partner Data: " + DateTime.Now.ToString());
                SyncChannelPartner(rootObject, service);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }
        private static void checkFCodePartner(IOrganizationService service, string cCode)
        {
            if (cCode.StartsWith("C"))
            {
                cCode = "F" + cCode;
            }
            Guid Partner = new Guid();
            Partner = Guid.Empty;
            QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, cCode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                Partner = Found.Entities[0].Id;
            }
            //return Partner;
        }
        private static void SyncChannelPartner(PartnerRootObject rootObject, IOrganizationService service)
        {
            string _KTOKD = string.Empty;
            Guid IfExists = Guid.Empty;
            string iKUNNR = string.Empty;
            Account iAccount = null;
            EntityReference entRefDistrict = null;
            int iDone = 0;
            int iTotal = rootObject.Results.Count;
            foreach (PartnerResult obj in rootObject.Results)
            {
                try
                {
                    iAccount = new Account();
                    if (obj.KUNNR.ToUpper() == "CVK0058")
                    {

                        _KTOKD = obj.KTOKD.ToString().Trim();

                        iKUNNR = obj.KUNNR;
                        Console.WriteLine("Channel Partner Code: " + iKUNNR);
                        if (iKUNNR.StartsWith("F"))
                            iKUNNR = iKUNNR.Substring(1);
                        IfExists = CheckIfPartnerExists(iKUNNR, service);

                        OptionSetValue _customerType = new OptionSetValue(_KTOKD == "0010" ? 1
                            : _KTOKD == "0020" ? 13 : _KTOKD == "0050" ? 14 : _KTOKD == "0045" ? 15
                            : _KTOKD == "0030" ? 16 : _KTOKD == "0031" ? 17 : _KTOKD == "0056" ? 6
                            : _KTOKD == "0065" ? 9 : _KTOKD == "0055" ? 6 : 12);
                        //if (_customerType.Value != 12)
                        //{
                        if (obj.delete_flag != "X")
                        {
                            iAccount.hil_InWarrantyCustomerSAPCode = obj.KUNNR;
                            iAccount.hil_OutWarrantyCustomerSAPCode = iKUNNR;
                            iAccount.Name = obj.VTXTM;
                            iAccount.EMailAddress1 = obj.SMTP_ADDR;
                            iAccount.Telephone1 = obj.MOB_NUMBER;
                            iAccount.Address1_Line1 = obj.STREET;
                            iAccount.Address1_Line2 = obj.STR_SUPPL3;
                            iAccount.Address1_Line3 = obj.ADDRESS3;
                            iAccount.AccountNumber = iKUNNR;
                            iAccount.hil_StagingPinUniqueKey = obj.KTOKD;
                            if (obj.Mtimestamp == null)
                                iAccount["hil_mdmtimestamp"] = ConvertToDateTime(obj.Ctimestamp);
                            else
                                iAccount["hil_mdmtimestamp"] = ConvertToDateTime(obj.Mtimestamp);

                            iAccount["address1_postofficebox"] = obj.GST_NO;
                            iAccount["hil_pan"] = obj.J_1IPANNO;

                            QueryExpression Query = new QueryExpression(hil_businessmapping.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet(true);
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            //Query.Criteria.AddCondition(new ConditionExpression("hil_stagingarea", ConditionOperator.Equal, obj.DM_AREA));
                            Query.Criteria.AddCondition(new ConditionExpression("hil_stagingpin", ConditionOperator.Equal, obj.dm_pin));
                            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                            EntityCollection Found = service.RetrieveMultiple(Query);
                            EntityReference salesOfficeRef = new EntityReference();
                            if (Found.Entities.Count > 0)
                            {
                                hil_businessmapping iBusMap = Found.Entities[0].ToEntity<hil_businessmapping>();
                                iAccount.hil_city = iBusMap.hil_city;
                                iAccount.hil_area = iBusMap.hil_area;
                                iAccount.hil_pincode = iBusMap.hil_pincode;
                                iAccount.hil_region = iBusMap.hil_region;
                                iAccount.hil_state = iBusMap.hil_state;
                                entRefDistrict = iBusMap.hil_district;
                                iAccount.hil_branch = iBusMap.hil_branch;
                                iAccount.hil_subterritory = iBusMap.hil_subterritory;
                                salesOfficeRef = iBusMap.hil_salesoffice;
                            }
                            if (obj.Longitude != null)
                                iAccount.Address1_Longitude = Double.Parse(obj.Longitude);
                            if (obj.Latitude != null)
                                iAccount.Address1_Latitude = Double.Parse(obj.Latitude);
                            iAccount.CustomerTypeCode = _customerType;
                            if (IfExists == Guid.Empty)
                            {
                                iAccount.hil_salesoffice = salesOfficeRef;
                                iAccount.hil_district = entRefDistrict;
                                // new OptionSetValue(_KTOKD == "0010" ? 1 : _KTOKD == "0020" ? 13 : _KTOKD == "0050" ? 14 : _KTOKD == "0045" ? 15 : _KTOKD == "0030" ? 16 : _KTOKD == "0031" ? 17 : _KTOKD == "0056" ? 6 : 9);
                                service.Create(iAccount);
                            }
                            else
                            {
                                if (CheckIfPartnerisFranchisee("F" + iKUNNR, service))
                                {
                                    iAccount.CustomerTypeCode = new OptionSetValue(6);
                                    iAccount.hil_InWarrantyCustomerSAPCode = "F" + iKUNNR;
                                }
                                iAccount.Id = IfExists;
                                service.Update(iAccount);
                            }
                        }
                        else if (obj.delete_flag == "X")
                        {
                            if (IfExists != Guid.Empty)
                            {
                                SetStateRequest setStateRequest = new SetStateRequest()
                                {
                                    EntityMoniker = new EntityReference
                                    {
                                        Id = IfExists,
                                        LogicalName = "account"
                                    },
                                    State = new OptionSetValue(1), //deactive
                                    Status = new OptionSetValue(2) //deactive
                                };
                                service.Execute(setStateRequest);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(obj.KUNNR.ToUpper());
                    }
                    //}
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error :" + ex.Message);
                }
                iDone = iDone + 1;
                Console.WriteLine("Record has been processed :" + iDone + "/" + iTotal);
            }
            Console.WriteLine(" TOTAL COUNT :" + iTotal.ToString());
        }
        public static DateTime? ConvertToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }
        public static string getTimeStamp(IOrganizationService service)
        {
            string _enquiryDatetime = "20210804000000";
            QueryExpression qsCType = new QueryExpression("account");
            qsCType.ColumnSet = new ColumnSet("hil_mdmtimestamp");
            qsCType.NoLock = true;
            qsCType.TopCount = 1;
            qsCType.AddOrder("hil_mdmtimestamp", OrderType.Descending);
            EntityCollection entCol = service.RetrieveMultiple(qsCType);
            if (entCol.Entities.Count > 0)
            {
                DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(300);
                if (_cTimeStamp.Year.ToString().PadLeft(4, '0') != "0001")
                    _enquiryDatetime = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
            }
            return _enquiryDatetime;
        }
        public static Guid CheckIfPartnerExists(string KUNNR, IOrganizationService service)
        {
            Guid Partner = new Guid();
            Partner = Guid.Empty;
            QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, KUNNR);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                Partner = Found.Entities[0].Id;
            }
            return Partner;
        }
        public static bool CheckIfPartnerisFranchisee(string KUNNR, IOrganizationService service)
        {
            bool Partner = false;

            QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_inwarrantycustomersapcode", ConditionOperator.Equal, KUNNR);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                Partner = true;
            }
            return Partner;
        }
        #endregion
        #region GET INTEGRATION URL
        public static Integration GetIntegration(IOrganizationService service, string RecName)
        {
            try
            {
                Integration intConf = new Integration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, RecName);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                intConf.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                intConf.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                intConf.passWord = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return intConf;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
                throw new Exception("Error : " + ex.Message);
            }
        }
        #endregion
    }
    public class Integration
    {
        public string uri { get; set; }
        public string userName { get; set; }
        public string passWord { get; set; }
    }
}
