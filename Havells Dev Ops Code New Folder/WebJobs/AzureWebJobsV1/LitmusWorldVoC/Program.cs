using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.Configuration;
using System.Net;
using System.ServiceModel.Description;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Query;
using System.Globalization;
using System.Threading;
using System.Text;
using RestSharp;
using Newtonsoft.Json;
using Microsoft.Xrm.Tooling.Connector;

namespace LitmusWorldVoC
{
    class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            try
            {
                _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
                if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
                {
                    GetFeedBackrequest(_service);
                }
            }
            catch (Exception ex)
            {
                Console.Write("ERROR!!! " + ex.Message);
            }
        }
        static void GetFeedBackrequest(IOrganizationService service)
        {
            /*
            1) 24hv_touchpoint - Post Breakdown Journey  --  {6560565A-3C0B-E911-A94E-000D3AF06CD4}
            2) az3f_touchpoint - Post Installation -- {E3129D79-3C0B-E911-A94E-000D3AF06CD4}
            3) b9r5_touchpoint - Post PMS -- {E2129D79-3C0B-E911-A94E-000D3AF06CD4}
            Job filters
            <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' /> // Closed
            <condition attribute='hil_isocr' operator='ne' value='1' />  does not OCR job
            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2022-04-15' />
            <condition attribute='hil_callertype' value='910590000' operator='ne'/>  Does not equal to Dealer
            */
            QueryExpression queryExp = new QueryExpression("hil_npssetup");
            queryExp.ColumnSet = new ColumnSet("hil_jobfiltercriteria");
            queryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            queryExp.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Consumer NPS"));

            EntityCollection entCol = _service.RetrieveMultiple(queryExp);
            if (entCol.Entities.Count > 0)
            {
                string _jobsfilterCriteria = entCol.Entities[0].GetAttributeValue<string>("hil_jobfiltercriteria");
                string strFetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_jobsextension'>
                <attribute name = 'createdon'/>
                <attribute name = 'hil_sentfornps'/>
                <filter type='and'>
                    <condition attribute='hil_sentfornps' value='1' operator='ne' />
                </filter>
                <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_jobs' link-type='inner' alias='ab'>
                <attribute name = 'hil_jobclosuredon'/>
                <attribute name = 'hil_tatcategory'/>
                <attribute name = 'hil_mobilenumber'/>
                <attribute name = 'hil_emailcustomer'/>
                <attribute name = 'hil_customerref'/>
                <attribute name = 'msdyn_substatus'/>
                <attribute name = 'msdyn_name'/>
                <attribute name = 'hil_callsubtype'/>
                <attribute name = 'hil_jobclosureon'/>
                <attribute name = 'hil_sparepartuse'/>
                <attribute name = 'hil_requesttype' />
                <attribute name = 'hil_sourceofjob'/>
                <attribute name = 'hil_salesoffice'/>
                <attribute name = 'hil_owneraccount'/>
                <attribute name = 'hil_region'/>
                <attribute name = 'hil_quantity'/>
                <attribute name = 'hil_productsubcategory'/>
                <attribute name = 'hil_productcatsubcatmapping'/>
                <attribute name = 'hil_productcategory'/>
                <attribute name = 'ownerid'/>
                <attribute name = 'hil_laborinwarranty'/>
                <attribute name = 'hil_kkgcode_sms'/>
                <attribute name = 'hil_jobclosemobile'/>
                <attribute name = 'hil_isocr'/>
                <attribute name = 'hil_isgascharged'/>
                <attribute name = 'hil_district'/>
                <attribute name = 'createdon'/>
                <attribute name = 'hil_countryclassification'/>
                <attribute name = 'hil_claimstatus'/>
                <attribute name = 'hil_channelpartnercategory'/>
                <attribute name = 'hil_callertype'/>
                <attribute name = 'hil_brand'/>
                <attribute name = 'msdyn_customerasset'/>
                <attribute name = 'hil_natureofcomplaint'/>
                <attribute name = 'hil_branch'/>
                <attribute name = 'hil_jobclosureon'/>
                <attribute name = 'hil_consumercategory'/>{0}
                <order attribute='msdyn_timeclosed' descending='false' />
                <link-entity name='hil_callsubtype' from='hil_callsubtypeid' to='hil_callsubtype' link-type='inner' alias='cs'>
                <attribute name = 'hil_npsprojectid'/>
                <filter type='and'>
                    <condition attribute='hil_npsprojectid' operator='not-null' />
                </filter>
                </link-entity>
                </link-entity>
                </entity>
                </fetch>";
                EntityCollection jobfeedbackCol = service.RetrieveMultiple(new FetchExpression(string.Format(strFetchXML, _jobsfilterCriteria)));
                if (jobfeedbackCol.Entities.Count > 0)
                {
                    int success = 0;
                    int TotalCount = jobfeedbackCol.Entities.Count;
                    int Fail = 0;

                    IntegrationConfiguration inconfig = GetIntegrationConfiguration(service);
                    //String sUrl = "https://p90ci.havells.com:50001/RESTAdapter/LitmusWorld/App/FeedbackRequest?IM_PROJECT=LitmusWorld";
                    String sUrl = inconfig.url;
                    String sUserName = inconfig.userName;
                    string sPassword = inconfig.password;
                    string _authInfo = sUserName + ":" + sPassword;

                    _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                    //string[] _Names = { "kuldeep Khare", "Nishank Jain", "PR Sajimon", "Lokesh Pandey", "SunilKumar Shukla", "Dheeraj Khanna", "Dhirendra Tripathy" };
                    //string[] _emails = { "kuldeep.khare@havells.com", "nishank.jain@havells.com", "pr.sajimon@havells.com", "Lokesh.Pandey@havells.com", "Sunilkumar.Shukla@havells.com", "DHEERAJ.KHANNA@HAVELLS.COM", "dhirendra.tripathy@havells.com" };
                    //string[] _mobiles = { "8285906486", "9910066033", "9958119947", "9811601235", "9650600776", "8800690751", "9739274180" };

                    try
                    {
                        FeedBackRequestRoot feedReq = new FeedBackRequestRoot();
                        Console.WriteLine("*******************NPS FeedBack Request**********************");
                        int i = 0;
                        foreach (Entity ent in jobfeedbackCol.Entities)
                        {
                            try
                            {
                                EntityReference _callSubType = ((EntityReference)ent.GetAttributeValue<AliasedValue>("ab.hil_callsubtype").Value);

                                if (ent.Contains("cs.hil_npsprojectid"))
                                {
                                    feedReq.app_id = ent.GetAttributeValue<AliasedValue>("cs.hil_npsprojectid").Value.ToString();
                                }
                                else
                                {
                                    continue;
                                }

                                feedReq.name = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_customerref").Value)).Name;
                                feedReq.customer_id = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_customerref").Value)).Id.ToString();
                                Entity custDetails = service.Retrieve("contact", new Guid(feedReq.customer_id), new ColumnSet("emailaddress1", "hil_premiumcustomer"));

                                if (custDetails.Contains("emailaddress1"))
                                {
                                    feedReq.userEmail = custDetails.GetAttributeValue<string>("emailaddress1");
                                }
                                else
                                {
                                    feedReq.userEmail = "";
                                }

                                feedReq.userPhone = (string)(ent.GetAttributeValue<AliasedValue>("ab.hil_mobilenumber").Value);

                                feedReq.tag_job_close_mobile = ((bool)(ent.GetAttributeValue<AliasedValue>("ab.hil_jobclosemobile").Value)).ToString();

                                if (custDetails.Contains("hil_premiumcustomer"))
                                {
                                    feedReq.tag_premium_customer = custDetails.GetAttributeValue<bool>("hil_premiumcustomer").ToString();
                                }
                                else
                                {
                                    feedReq.tag_premium_customer = "";
                                }

                                feedReq.tag_job_id = (string)ent.GetAttributeValue<AliasedValue>("ab.msdyn_name").Value;
                                feedReq.tag_work_done_on = ((DateTime)(ent.GetAttributeValue<AliasedValue>("ab.hil_jobclosuredon").Value)).Date.ToString();
                                if (ent.Contains("ab.hil_tatcategory"))
                                {
                                    feedReq.tag_tat_category = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_tatcategory").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_tat_category = "";
                                }
                                feedReq.tag_sub_status = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.msdyn_substatus").Value)).Name;
                                if (ent.Contains("ab.hil_callsubtype"))
                                {
                                    feedReq.tag_call_subtype = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_callsubtype").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_call_subtype = "";
                                }
                                if (ent.Contains("ab.hil_productcategory"))
                                {
                                    feedReq.tag_product_category = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_productcategory").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_product_category = "";
                                }
                                if (ent.Contains("ab.hil_productcatsubcatmapping"))
                                {
                                    feedReq.tag_product_subcategory = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_productcatsubcatmapping").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_product_subcategory = "";
                                }
                                if (ent.Contains("ab.msdyn_customerasset"))
                                {
                                    feedReq.tag_associated_customer_product = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.msdyn_customerasset").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_associated_customer_product = "";
                                }
                                if (ent.Contains("ab.hil_natureofcomplaint"))
                                {
                                    feedReq.tag_nature_of_complaint = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_natureofcomplaint").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_nature_of_complaint = "";
                                }
                                if (ent.Contains("ab.ownerid"))
                                {
                                    feedReq.tag_owner_technician = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.ownerid").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_owner_technician = "";
                                }
                                if (ent.Contains("ab.hil_branch"))
                                {
                                    feedReq.tag_private_group_hierarchy_0_id = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_branch").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_private_group_hierarchy_0_id = "";
                                }
                                if (ent.Contains("ab.hil_district"))
                                {
                                    feedReq.tag_district = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_district").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_district = "";
                                }
                                feedReq.tag_created_on = ((DateTime)(ent.GetAttributeValue<AliasedValue>("ab.createdon").Value)).Date.ToString();
                                if (ent.Contains("ab.hil_brand"))
                                {
                                    feedReq.tag_brand = ent.FormattedValues["ab.hil_brand"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_brand = "";

                                }
                                if (ent.Contains("ab.hil_owneraccount"))
                                {
                                    feedReq.tag_related_franchisee = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_owneraccount").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_related_franchisee = "";

                                }
                                if (ent.Contains("ab.hil_jobclosureon"))
                                {
                                    feedReq.tag_ticket_closed_on = ((DateTime)(ent.GetAttributeValue<AliasedValue>("ab.hil_jobclosureon").Value)).Date.ToString();
                                }
                                else
                                {
                                    feedReq.tag_ticket_closed_on = "";
                                }
                                if (ent.Contains("ab.hil_quantity"))
                                {
                                    feedReq.tag_quantity = ((int)(ent.GetAttributeValue<AliasedValue>("ab.hil_quantity").Value)).ToString();
                                }
                                else
                                {
                                    feedReq.tag_quantity = "";
                                }
                                if (ent.Contains("ab.hil_isocr"))
                                {
                                    feedReq.tag_is_ocr = ((bool)(ent.GetAttributeValue<AliasedValue>("ab.hil_isocr").Value)).ToString();
                                }
                                else
                                {
                                    feedReq.tag_is_ocr = "";
                                }
                                if (ent.Contains("ab.hil_kkgcode_sms"))
                                {
                                    feedReq.tag_kkg_used = ent.FormattedValues["ab.hil_kkgcode_sms"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_kkg_used = "";
                                }
                                if (ent.Contains("ab.hil_claimstatus"))
                                {
                                    feedReq.tag_claim_status = ent.FormattedValues["ab.hil_claimstatus"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_claim_status = "";
                                }
                                if (ent.Contains("ab.hil_countryclassification"))
                                {
                                    feedReq.tag_country_classification = ent.FormattedValues["ab.hil_countryclassification"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_country_classification = "";
                                }
                                if (ent.Contains("ab.hil_isgascharged"))
                                {
                                    feedReq.tag_is_gas_charged = ((bool)(ent.GetAttributeValue<AliasedValue>("ab.hil_isgascharged").Value)).ToString();
                                }
                                else
                                {
                                    feedReq.tag_is_gas_charged = "";
                                }
                                if (ent.Contains("ab.hil_sparepartuse"))
                                {
                                    feedReq.tag_spare_part_used = ((bool)(ent.GetAttributeValue<AliasedValue>("ab.hil_sparepartuse").Value)).ToString();
                                }
                                else
                                {
                                    feedReq.tag_spare_part_used = "";
                                }
                                if (ent.Contains("ab.hil_callertype"))
                                {
                                    feedReq.tag_caller_type = ent.FormattedValues["ab.hil_callertype"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_caller_type = "";
                                }
                                if (ent.Contains("ab.hil_channelpartnercategory"))
                                {
                                    feedReq.tag_channel_partner_category = ent.FormattedValues["ab.hil_channelpartnercategory"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_channel_partner_category = "";
                                }
                                if (ent.Contains("ab.hil_consumercategory"))
                                {
                                    feedReq.tag_consumer_category = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_consumercategory").Value)).ToString();
                                }
                                else
                                {
                                    feedReq.tag_consumer_category = "";
                                }
                                if (ent.Contains("ab.hil_laborinwarranty"))
                                {
                                    feedReq.tag_labor_in_warranty = ((bool)(ent.GetAttributeValue<AliasedValue>("ab.hil_laborinwarranty").Value)).ToString();
                                }
                                else
                                {
                                    feedReq.tag_labor_in_warranty = "";
                                }
                                if (ent.Contains("ab.hil_salesoffice"))
                                {
                                    feedReq.tag_sales_office = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_salesoffice").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_sales_office = "";
                                }
                                if (ent.Contains("ab.hil_sourceofjob"))
                                {
                                    feedReq.tag_source_of_job = ent.FormattedValues["ab.hil_sourceofjob"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_source_of_job = "";
                                }
                                if (ent.Contains("ab.hil_requesttype"))
                                {
                                    feedReq.tag_source_of_job_closure = ent.FormattedValues["ab.hil_requesttype"].ToString();
                                }
                                else
                                {
                                    feedReq.tag_source_of_job_closure = "";
                                }
                                if (ent.Contains("ab.hil_region"))
                                {
                                    feedReq.tag_region = ((EntityReference)(ent.GetAttributeValue<AliasedValue>("ab.hil_region").Value)).Name;
                                }
                                else
                                {
                                    feedReq.tag_region = "";
                                }
                                string data = JsonConvert.SerializeObject(feedReq);
                                var client = new RestClient(sUrl);
                                client.Timeout = -1;
                                var request = new RestRequest(Method.POST);
                                request.AddHeader("Content-Type", "application/json");
                                request.AddParameter("application/json", data, ParameterType.RequestBody);
                                request.AddHeader("Authorization", "Basic " + _authInfo);
                                IRestResponse response = client.Execute(request);
                                Console.WriteLine("Data Fetched");
                                ResponseRoot rRoot = new ResponseRoot();
                                rRoot = JsonConvert.DeserializeObject<ResponseRoot>(response.Content);
                                if (rRoot.data.error_message == null)
                                {
                                    Entity jobextension = new Entity("hil_jobsextension");
                                    jobextension.Id = ent.Id;
                                    jobextension["hil_sentfornps"] = true;
                                    service.Update(jobextension);
                                    success += 1;
                                    Console.WriteLine("NPS Success Count " + success + "/" + TotalCount);
                                }
                                else
                                {
                                    Fail += 1;
                                    Console.WriteLine("NPS request failed: " + Fail);
                                }
                                i++;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error Message " + ex.Message);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error Message " + ex.Message);
                        Console.WriteLine("NPS Success Count " + success + "/" + TotalCount);
                        Console.WriteLine("NPS Fail Count " + Fail);
                    }
                }
            }
        }
        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion

        private static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service)
        {
            try
            {
                IntegrationConfiguration inconfig = new IntegrationConfiguration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "NPSFeedbackRequestAPI");
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return inconfig;
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
    }

    #region DTOs
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }
    public class ResponseData
    {
        public string error_message { get; set; }
        public int code { get; set; }
    }
    public class ResponseRoot
    {
        public int code { get; set; }
        public ResponseData data { get; set; }
    }
    public class FeedBackRequestRoot
    {
        public string app_id { get; set; }
        public string userPhone { get; set; }
        public string userEmail { get; set; }
        public string name { get; set; }
        public string customer_id { get; set; }
        public string tag_job_id { get; set; }
        public string tag_work_done_on { get; set; }
        public string tag_tat_category { get; set; }
        public string tag_sub_status { get; set; }
        public string tag_call_subtype { get; set; }
        public string tag_product_category { get; set; }
        public string tag_product_subcategory { get; set; }
        public string tag_associated_customer_product { get; set; }
        public string tag_nature_of_complaint { get; set; }
        public string tag_owner_technician { get; set; }
        public string tag_private_group_hierarchy_0_id { get; set; }
        public string tag_district { get; set; }
        public string tag_created_on { get; set; }
        public string tag_brand { get; set; }
        public string tag_related_franchisee { get; set; }
        public string tag_ticket_closed_on { get; set; }
        public string tag_quantity { get; set; }
        public string tag_is_ocr { get; set; }
        public string tag_job_close_mobile { get; set; }
        public string tag_kkg_used { get; set; }
        public string tag_claim_status { get; set; }
        public string tag_country_classification { get; set; }
        public string tag_is_gas_charged { get; set; }
        public string tag_spare_part_used { get; set; }
        public string tag_caller_type { get; set; }
        public string tag_channel_partner_category { get; set; }
        public string tag_consumer_category { get; set; }
        public string tag_labor_in_warranty { get; set; }
        public string tag_sales_office { get; set; }
        public string tag_source_of_job { get; set; }
        public string tag_source_of_job_closure { get; set; }
        public string tag_region { get; set; }
        public string tag_premium_customer { get; set; }
    }
    #endregion
}