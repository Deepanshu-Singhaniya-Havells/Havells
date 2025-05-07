using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System.Text;

namespace ChannelPartnerSync
{
    public class ClsPartner
    {
        private readonly ServiceClient service;
        public ClsPartner(ServiceClient _service)
        {
            service = _service;
        }
        public void GetChannelPartnerData(string? channelPartnerCode)
        {
            try
            {
                Integration intConf = GetIntegration("Tier 1 Customer");
                string uri = intConf.uri;
                string authInfo = intConf.userName + ":" + intConf.passWord;
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                if (!string.IsNullOrWhiteSpace(channelPartnerCode))
                {
                    uri = $"{uri}ChanelPartnerCode={channelPartnerCode}";
                }
                else
                {
                    string enquiryDate = DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "000000"; // Yesterday's date
                    string todate = DateTime.Now.ToString("yyyyMMdd") + "000000"; // Today's date
                    uri = $"{uri}enquiryDate={enquiryDate}&ToDate={todate}";
                }
                Console.WriteLine("WebJob Execution for URL: " + uri);
                var client = new RestClient();
                var request = new RestRequest(uri, Method.Post);
                request.AddHeader("Authorization", "Basic " + authInfo);
                var response = client.Execute(request);
                PartnerRootObject rootObject = JsonConvert.DeserializeObject<PartnerRootObject>(response.Content);
                SyncChannelPartner(rootObject);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }
        private void SyncChannelPartner(PartnerRootObject rootObject)
        {
            string _KTOKD = string.Empty;
            Guid IfExists = Guid.Empty;
            string iKUNNR = string.Empty;
            EntityReference entRefDistrict = null;
            int iDone = 0;
            int iTotal = rootObject.Results.Count;
            Console.WriteLine($"Total Channel Partner Count: {rootObject.Results.Count}");
            foreach (PartnerResult obj in rootObject.Results)
            {
                try
                {
                    Entity iAccount = new Entity("account");
                    _KTOKD = obj.KTOKD.ToString().Trim();

                    iKUNNR = obj.KUNNR;
                    Console.WriteLine("Channel Partner Code: " + iKUNNR);
                    if (iKUNNR.StartsWith("F"))
                        iKUNNR = iKUNNR.Substring(1);
                    IfExists = CheckIfPartnerExists(iKUNNR);

                    OptionSetValue _customerType = new OptionSetValue(_KTOKD == "0010" ? 1
                        : _KTOKD == "0020" ? 13 : _KTOKD == "0050" ? 14 : _KTOKD == "0045" ? 15
                        : _KTOKD == "0030" ? 16 : _KTOKD == "0031" ? 17 : _KTOKD == "0056" ? 6
                        : _KTOKD == "0065" ? 9 : _KTOKD == "0055" ? 6 : 12);
                    if (obj.delete_flag != "X")
                    {
                        iAccount["hil_inwarrantycustomersapcode"] = obj.KUNNR;
                        iAccount["hil_outwarrantycustomersapcode"] = iKUNNR;
                        iAccount["name"] = obj.VTXTM;
                        iAccount["emailaddress1"] = obj.SMTP_ADDR;
                        iAccount["telephone1"] = MobileNumberLength(obj.MOB_NUMBER);
                        iAccount["address1_line1"] = obj.STREET;
                        iAccount["address1_line2"] = obj.STR_SUPPL3;
                        iAccount["address1_line3"] = obj.ADDRESS3;
                        iAccount["accountnumber"] = iKUNNR;
                        iAccount["hil_stagingpinuniquekey"] = obj.KTOKD;
                        if (obj.Mtimestamp == null)
                            iAccount["hil_mdmtimestamp"] = ConvertToDateTime(obj.Ctimestamp);
                        else
                            iAccount["hil_mdmtimestamp"] = ConvertToDateTime(obj.Mtimestamp);

                        iAccount["address1_postofficebox"] = obj.GST_NO;
                        iAccount["hil_pan"] = obj.J_1IPANNO;

                        QueryExpression Query = new QueryExpression("hil_businessmapping");
                        Query.ColumnSet = new ColumnSet(true);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition(new ConditionExpression("hil_stagingpin", ConditionOperator.Equal, obj.dm_pin));
                        Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        EntityReference salesOfficeRef = new EntityReference();
                        if (Found.Entities.Count > 0)
                        {
                            iAccount["hil_city"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_city"); //iBusMap.hil_city;
                            iAccount["hil_area"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_area");// iBusMap.hil_area;
                            iAccount["hil_pincode"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_pincode"); // iBusMap.hil_pincode;
                            iAccount["hil_region"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_region"); //iBusMap.hil_region;
                            iAccount["hil_state"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_state"); //iBusMap.hil_state;
                            entRefDistrict = Found.Entities[0].GetAttributeValue<EntityReference>("hil_district"); //iBusMap.hil_district;
                            iAccount["hil_branch"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_branch"); // iBusMap.hil_branch;
                            iAccount["hil_subterritory"] = Found.Entities[0].GetAttributeValue<EntityReference>("hil_subterritory"); // iBusMap.hil_subterritory;
                            salesOfficeRef = Found.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice"); //iBusMap.hil_salesoffice;
                        }
                        if (obj.Longitude != null)
                            iAccount["address1_longitude"] = Double.Parse(obj.Longitude);
                        if (obj.Latitude != null)
                            iAccount["address1_latitude"] = Double.Parse(obj.Latitude);
                        iAccount["customertypecode"] = _customerType;
                        if (IfExists == Guid.Empty)
                        {
                            iAccount["hil_salesoffice"] = salesOfficeRef;
                            iAccount["hil_district"] = entRefDistrict;
                            service.Create(iAccount);
                        }
                        else
                        {
                            if (CheckIfPartnerisFranchisee("F" + iKUNNR))
                            {
                                iAccount["customertypecode"] = new OptionSetValue(6);
                                iAccount["hil_inwarrantycustomersapcode"] = "F" + iKUNNR;
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
                    else
                    {
                        Console.WriteLine(obj.KUNNR.ToUpper());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : " + ex.Message);
                }
                iDone = iDone + 1;
                Console.WriteLine("Record has been processed : " + iDone + "/" + iTotal);
            }
            Console.WriteLine(" TOTAL COUNT : " + iTotal.ToString());
        }
        public DateTime? ConvertToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }
        public Guid CheckIfPartnerExists(string KUNNR)
        {
            Guid Partner = new Guid();
            Partner = Guid.Empty;
            QueryExpression Query = new QueryExpression("account");
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
        public bool CheckIfPartnerisFranchisee(string KUNNR)
        {
            bool Partner = false;
            QueryExpression Query = new QueryExpression("account");
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
        private string MobileNumberLength(string MOB_NUMBER)
        {
            string finalMobNumber = MOB_NUMBER;
            if (MOB_NUMBER.Length > 10)
            {
                finalMobNumber = MOB_NUMBER.Substring(MOB_NUMBER.Length - 10, 10);
            }
            return finalMobNumber;
        }
        public Integration GetIntegration(string RecName)
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
    }
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
    public class Integration
    {
        public string uri { get; set; }
        public string userName { get; set; }
        public string passWord { get; set; }
    }
}
