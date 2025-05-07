using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using Havells.Crm.CommonLibrary;
using System.IO;
using System.Net;
using Microsoft.Crm.Sdk.Messages;
using System.Security.Policy;

namespace Havells.CRM.WhatsAppPromotions
{
    public class Program : AzureWebJobsLogs
    {
        #region Global Varialble declaration
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        public static IOrganizationService service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        public static string containerName = "WebJobLogs-WhatsAppPromotions";
        public static string fileName = "";
        public static string msg = null;
        public static string header = null;
        public static string blobUrl = null;
        #endregion
        static Program()
        {
            service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            //string path = DateTime.Now.ToString("ddMMMyyyy") + "/";
            //fileName = path + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + ".csv";
            //header = "Mobile Number,Serial Number, Job Id, Mode Of Comunication,Template,URL,Payload,Response";
            //blobUrl = CreateOrUpdateLogs(containerName, fileName, msg, header);
        }
        static void Main(string[] args)
        {
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)service).IsReady)
            {
                try
                {
                    //#region New Buyer +14 days (WA)
                    //Campaigns.Campaign_NewBuyer("New Buyer +14 days (AC-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.AC_AMC_14D, -14, ProductCategoryGUID.AirConditioner);
                    //Campaigns.Campaign_NewBuyer("New Buyer +14 days (WP-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.WP_AMC_14D, -14, ProductCategoryGUID.WaterPurifier);
                    //#endregion

                    //#region New Buyer +30 days (SMS)
                    //Campaigns.Campaign_NewBuyer("New Buyer + 30 days(AC-SMS)", ModeOfCommunication.SMS, "1107168654225056811", -30, ProductCategoryGUID.AirConditioner);
                    //Campaigns.Campaign_NewBuyer("New Buyer + 30 days(WP-SMS)", ModeOfCommunication.SMS, "1107169702429550488", -30, ProductCategoryGUID.WaterPurifier);
                    //#endregion

                    //#region Breakdown Job Closure +7 Days (WA)
                    //Campaigns.Campaign_OutWarrantyBreakdownJobClosure("Breakdown Job Closure +7 Days (AC-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.AC_AMC_7D, -7, ProductCategoryGUID.AirConditioner);
                    //Campaigns.Campaign_OutWarrantyBreakdownJobClosure("Breakdown Job Closure +7 Days (WP-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.WP_AMC_7D, -7, ProductCategoryGUID.WaterPurifier);
                    //#endregion

                    //#region Warranty Expiry -30 Days (WA)
                    //Campaigns.Campaign_WarrantyExpireNear("Warranty Expiry -30 Days (AC-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.AC_AMC_30D, 30, ProductCategoryGUID.AirConditioner);
                    //Campaigns.Campaign_WarrantyExpireNear("Warranty Expiry -30 Days (WP-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.WP_AMC_30D, 30, ProductCategoryGUID.WaterPurifier);
                    //#endregion

                    SendCampaign.RetriveCamapign();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error!!! " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Error in establishing connection with Dynamics.");
            }
        }
        static protected void UnitWarrantyRetrive(ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string ProductCategoryGUID)
        {
            int done = 0;
            int error = 0;
            int skip = 0;
            int totalCount = 0;
            logs _logs = new logs();
            try
            {
                string dateString = DateTime.Today.AddDays(dayDif).Year.ToString().PadLeft(4, '0') + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                while (moreRecord)
                {
                    string fetchUnitwarranty = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                                 <entity name='hil_unitwarranty'>
                                                    <attribute name='hil_name'/>
                                                    <attribute name='createdon'/>
                                                    <attribute name='hil_warrantytemplate'/>
                                                    <attribute name='hil_warrantystartdate'/>
                                                    <attribute name='hil_warrantyenddate'/>
                                                    <attribute name='hil_producttype'/>
                                                    <attribute name='hil_customerasset'/>
                                                    <attribute name='hil_unitwarrantyid'/>
                                                    <order attribute='hil_name' descending='false'/>
                                                    <filter type='and'>
                                                        <condition attribute='hil_warrantyenddate' operator='on' value='{dateString}'/>
                                                    </filter>
                                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='ac'>
                                                        <filter type='and'>
                                                            <condition attribute='hil_type' operator='in'>
                                                                <value>3</value>
                                                                <value>1</value>
                                                            </condition>
                                                        </filter>
                                                    </link-entity>
                                                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='hil_customerasset' link-type='inner' alias='ad'>
                                                        <attribute name='hil_customer'/>
                                                        <attribute name='msdyn_name'/>
                                                        <attribute name='hil_productsubcategorymapping'/>
                                                        <attribute name='msdyn_product'/>
                                                        <filter type='and'>
                                                            <condition attribute='hil_productcategory' operator='eq' uiname='HAVELLS AQUA' uitype='product' value='{ProductCategoryGUID}'/>
                                                        </filter>
                                                    </link-entity>
                                                  </entity>
                                                </fetch>";

                    EntityCollection warrantyColl = service.RetrieveMultiple(new FetchExpression(fetchUnitwarranty));
                    if (warrantyColl.Entities.Count > 0)

                    {
                        cuatomerAssetColl.Entities.AddRange(warrantyColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }

                string customerName = "Customer";
                string customerMobileNumber = string.Empty;
                string productsubcategory = string.Empty;
                string productmodel = string.Empty;
                string serialNumber = string.Empty;
                string installation_date = string.Empty;

                Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
                totalCount = cuatomerAssetColl.Entities.Count;
                foreach (Entity job in cuatomerAssetColl.Entities)
                {

                    customerName = "Customer";
                    customerMobileNumber = string.Empty;
                    productmodel = string.Empty;
                    productsubcategory = string.Empty;
                    serialNumber = string.Empty;
                    installation_date = string.Empty;
                    _logs = new logs();
                    try
                    {
                        string dateStringNew = DateTime.Today.AddDays(dayDif + 2).Year.ToString().PadLeft(4, '0') + "-" + DateTime.Today.AddDays(dayDif + 2).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif + 2).Day.ToString().PadLeft(2, '0');

                        string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                <entity name='hil_unitwarranty'>
                                                    <attribute name='hil_name'/>
                                                    <attribute name='createdon'/>
                                                    <attribute name='hil_warrantytemplate'/>
                                                    <attribute name='hil_warrantystartdate'/>
                                                    <attribute name='hil_warrantyenddate'/>
                                                    <attribute name='hil_producttype'/>
                                                    <attribute name='hil_customerasset'/>
                                                    <attribute name='hil_unitwarrantyid'/>
                                                    <order attribute='hil_name' descending='false'/>
                                                    <filter type='and'>
                                                        <condition attribute='hil_customerasset' operator='eq' value='{job.GetAttributeValue<EntityReference>("hil_customerasset").Id}'/>
                                                        <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{dateStringNew}'/>
                                                    </filter>
                                                </entity>
                                            </fetch>";

                        EntityCollection unitwarrColl = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (unitwarrColl.Entities.Count > 0)
                        {
                            skip++;
                            // Console.WriteLine("Skiped " + skip + "/" + cuatomerAssetColl.Entities.Count + " of searial Number " + job.GetAttributeValue<EntityReference>("hil_customerasset").Name);

                        }
                        else
                        {
                            customerName = job.Contains("ad.hil_customer") ? ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_customer").Value).Name.ToString() : "";

                            customerMobileNumber = service.Retrieve("contact", ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_customer").Value).Id,
                                new ColumnSet("mobilephone")).GetAttributeValue<string>("mobilephone");


                            ///////Customer Mobile Number HardCode
                            // customerMobileNumber = "8005084995";


                            productsubcategory = job.Contains("ad.hil_productsubcategorymapping") ? ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.hil_productsubcategorymapping").Value).Name.ToString() : "";
                            productmodel = job.Contains("ad.msdyn_product") ? ((EntityReference)job.GetAttributeValue<AliasedValue>("ad.msdyn_product").Value).Name.ToString() : "";
                            serialNumber = job.Contains("ad.msdyn_name") ? job.GetAttributeValue<AliasedValue>("ad.msdyn_name").Value.ToString() : "";
                            installation_date = job.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");

                            _logs.SerialNumber = serialNumber;
                            _logs.MobileNumber = customerMobileNumber;
                            _logs.JobId = "-";
                            _logs.Template = templateName;
                            _logs.ModeOfComunication = _ModeOfComm == ModeOfCommunication.Whatsapp ? "Whats App" : "SMS";


                            if (_ModeOfComm == ModeOfCommunication.Whatsapp)
                            {
                                UserDetails userDetails = new UserDetails();
                                Traits objtraits = new Traits();
                                objtraits.name = customerName;

                                userDetails.phoneNumber = customerMobileNumber;
                                userDetails.countryCode = "+91";
                                userDetails.traits = objtraits;
                                userDetails.tags = new List<object>();

                                UpdateWhatsAppUser(userDetails, _logs);

                                CampaignDetails campaignDetails = new CampaignDetails();
                                campaignDetails.phoneNumber = customerMobileNumber;
                                campaignDetails.countryCode = "+91";
                                campaignDetails.@event = templateName;

                                Traits campaignTraits = new Traits();
                                campaignTraits.expire_on = " ";
                                if (templateName.Contains("_WP_"))// == "5D_WP_AMC"|| templateName == "21D_WP_AMC")
                                {
                                    campaignTraits.prd_cat = "Havells Aqua";
                                }
                                if (templateName.Contains("_AC_"))// == "5D_AC_AMC" || templateName == "21D_AC_AMC")
                                {
                                    campaignTraits.prd_cat = "Lloyd Air Conditioner";
                                }
                                campaignTraits.product_model = productsubcategory;
                                campaignTraits.product_serial_number = serialNumber;
                                campaignTraits.registration_date = installation_date;

                                campaignDetails.traits = campaignTraits;
                                SendMessage(campaignDetails, _logs);
                                done++;
                                Console.WriteLine("Executed Whatsapp Campaign# " + templateName + " -> " + done.ToString() + "/" + cuatomerAssetColl.Entities.Count + " to " + job.GetAttributeValue<EntityReference>("hil_customerasset").Name);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        error++;
                        Console.WriteLine("***** " + error + "/" + cuatomerAssetColl.Entities.Count + "Error Occured " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
            }
            Console.WriteLine("===================================================================");
            Console.WriteLine("Summary");
            Console.WriteLine("Template : " + templateName);
            Console.WriteLine("Total Recourd : " + totalCount);
            Console.WriteLine("Skiped : " + skip);
            Console.WriteLine("Send : " + done);
            Console.WriteLine("===================================================================");
            SendEmailOnCampaignCompleation(service, templateName, done, _ModeOfComm == ModeOfCommunication.SMS ? "SMS" : "What's App");
        }
        static protected void OutWarrantyBreakDownJobs(ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string ProductCategoryGUID, string Callsubtype)
        {
            int done = 0;
            int error = 0;
            int skip = 0;
            int totalCount = 0;
            logs _logs = new logs();
            try
            {
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                while (moreRecord)
                {
                    string fetchoutwarranty = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                                   <entity name='msdyn_workorder'>
                                                       <attribute name='msdyn_name'/>
                                                       <attribute name='createdon'/>
                                                       <attribute name='hil_productsubcategory'/>
                                                       <attribute name='hil_customerref'/>
                                                        <attribute name='hil_callsubtype'/>
                                                        <attribute name='msdyn_workorderid'/>
                                                        <attribute name='msdyn_timeclosed'/>
                                                        <attribute name='msdyn_customerasset'/>
                                                        <attribute name='hil_laborinwarranty'/>
                                                        <order attribute='msdyn_name' descending='false'/>
                                                        <filter type='and'>
                                                            <condition attribute='msdyn_substatus' operator='eq' value='{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}'/>
                                                            <condition attribute='hil_isocr' operator='ne' value='1'/>
                                                            <condition attribute='msdyn_timeclosed' operator='on' value='{dateString}'/>
                                                            <condition attribute='hil_productcategory' operator='in'>
                                                                <value>{ProductCategoryGUID}</value>
                                                            </condition>
                                                            <condition attribute='hil_callsubtype' operator='eq' value='{Callsubtype}'/>
                                                        </filter>
                                                        <link-entity name=""msdyn_customerasset"" from=""msdyn_customerassetid"" to=""msdyn_customerasset"" visible=""false"" link-type=""outer"" alias=""a"">
                                                            <attribute name=""msdyn_product""/>
                                                        </link-entity>
                                                    </entity>
                                                </fetch>";

                    EntityCollection outWarrantyColl = service.RetrieveMultiple(new FetchExpression(fetchoutwarranty));
                    if (outWarrantyColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(outWarrantyColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }
                string customerName = "Customer";
                string customerMobileNumber = string.Empty;
                string productsubcategory = string.Empty;
                string productmodel = string.Empty;
                string serialNumber = string.Empty;
                string installation_date = string.Empty;
                Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
                totalCount = cuatomerAssetColl.Entities.Count;
                foreach (Entity job in cuatomerAssetColl.Entities)
                {

                    customerName = "Customer";
                    customerMobileNumber = string.Empty;
                    productsubcategory = string.Empty;
                    productmodel = string.Empty;
                    serialNumber = string.Empty;
                    installation_date = string.Empty;
                    try
                    {
                        string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                <entity name='hil_unitwarranty'>
                                                    <attribute name='hil_name'/>
                                                    <attribute name='createdon'/>
                                                    <attribute name='hil_warrantytemplate'/>
                                                    <attribute name='hil_warrantystartdate'/>
                                                    <attribute name='hil_warrantyenddate'/>
                                                    <attribute name='hil_producttype'/>
                                                    <attribute name='hil_customerasset'/>
                                                    <attribute name='hil_unitwarrantyid'/>
                                                    <order attribute='hil_name' descending='false'/>
                                                    <filter type='and'>
                                                        <condition attribute='hil_customerasset' operator='eq' value='{job.GetAttributeValue<EntityReference>("msdyn_customerasset").Id}'/>
                                                        <condition attribute='hil_warrantyenddate' operator='on-or-after' value='{dateString}'/>
                                                    </filter>
                                                </entity>
                                            </fetch>";

                        EntityCollection unitwarrColl = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (unitwarrColl.Entities.Count > 0)
                        {
                            skip++;
                            //Console.WriteLine("Skiped " + skip + "/" + cuatomerAssetColl.Entities.Count + " of searial Number " + job.GetAttributeValue<EntityReference>("msdyn_customerasset").Name);

                        }
                        else
                        {
                            _logs = new logs();
                            customerName = job.Contains("hil_customerref") ? job.GetAttributeValue<EntityReference>("hil_customerref").Name : "";
                            customerMobileNumber = service.Retrieve("contact", job.GetAttributeValue<EntityReference>("hil_customerref").Id,
                                new ColumnSet("mobilephone")).GetAttributeValue<string>("mobilephone");

                            ///////Customer Mobile Number HardCode
                            // customerMobileNumber = "8005084995";

                            productsubcategory = job.GetAttributeValue<EntityReference>("hil_productsubcategory").Name.ToString();
                            productmodel = ((EntityReference)job.GetAttributeValue<AliasedValue>("a.msdyn_product").Value).Name;
                            serialNumber = job.GetAttributeValue<String>("msdyn_name").ToString();
                            installation_date = job.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");
                            _logs.SerialNumber = serialNumber;
                            _logs.MobileNumber = customerMobileNumber;
                            _logs.JobId = "-";
                            _logs.Template = templateName;
                            _logs.ModeOfComunication = _ModeOfComm == ModeOfCommunication.Whatsapp ? "Whats App" : "SMS";
                            if (_ModeOfComm == ModeOfCommunication.Whatsapp)
                            {
                                UserDetails userDetails = new UserDetails();
                                Traits objtraits = new Traits();
                                objtraits.name = customerName;

                                userDetails.phoneNumber = customerMobileNumber;
                                userDetails.countryCode = "+91";
                                userDetails.traits = objtraits;
                                userDetails.tags = new List<object>();

                                UpdateWhatsAppUser(userDetails, _logs);

                                CampaignDetails campaignDetails = new CampaignDetails();
                                campaignDetails.phoneNumber = customerMobileNumber;
                                campaignDetails.countryCode = "+91";
                                campaignDetails.@event = templateName;

                                Traits campaignTraits = new Traits();
                                campaignTraits.expire_on = " ";
                                if (templateName.Contains("_WP_"))// == "5D_WP_AMC"|| templateName == "21D_WP_AMC")
                                {
                                    campaignTraits.prd_cat = "Havells Aqua";
                                }
                                if (templateName.Contains("_AC_"))// == "5D_AC_AMC" || templateName == "21D_AC_AMC")
                                {
                                    campaignTraits.prd_cat = "Lloyd Air Conditioner";
                                }
                                campaignTraits.product_model = productsubcategory;
                                campaignTraits.product_serial_number = serialNumber;
                                campaignTraits.registration_date = installation_date;

                                campaignDetails.traits = campaignTraits;
                                SendMessage(campaignDetails, _logs);
                                done++;
                                Console.WriteLine("Executed Whatsapp Campaign# " + templateName + " -> " + done.ToString() + "/" + cuatomerAssetColl.Entities.Count + " to " + customerMobileNumber);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        error++;
                        Console.WriteLine("***** " + error + "/" + cuatomerAssetColl.Entities.Count + "Error Occured " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
            }
            Console.WriteLine("===================================================================");
            Console.WriteLine("Summary");
            Console.WriteLine("Template : " + templateName);
            Console.WriteLine("Total Recourd : " + totalCount);
            Console.WriteLine("Skiped : " + skip);
            Console.WriteLine("Send : " + done);
            Console.WriteLine("===================================================================");
            SendEmailOnCampaignCompleation(service, templateName, done, _ModeOfComm == ModeOfCommunication.SMS ? "SMS" : "What's App");
        }
        static protected void retriveJobs(ModeOfCommunication _ModeOfComm, string templateName, int dayDif, string ProductCategoryGUID)
        {
            int done = 0;
            int error = 0;
            int skip = 0;
            int sucess = 0;
            int totalCount = 0;
            try
            {
                string dateString = DateTime.Today.AddDays(dayDif).Year + "-" + DateTime.Today.AddDays(dayDif).Month.ToString().PadLeft(2, '0') + "-" + DateTime.Today.AddDays(dayDif).Day.ToString().PadLeft(2, '0');
                int pageNumber = 1;
                EntityCollection cuatomerAssetColl = new EntityCollection();
                bool moreRecord = true;
                while (moreRecord)
                {
                    string fetchCUstomerAsset = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                    <entity name=""msdyn_customerasset"">
                    <attribute name=""msdyn_customerassetid"" />
                    <attribute name=""msdyn_name"" />
                    <attribute name=""hil_productcategory"" />
                    <attribute name=""hil_productsubcategorymapping"" />
                    <attribute name=""msdyn_product"" />
                    <attribute name=""hil_customer"" />
                    <attribute name='createdon' />
                    <order attribute=""msdyn_name"" descending=""false"" />
                    <filter type=""and"">
                    <condition attribute=""createdon"" operator=""on"" value=""{dateString}"" />
                    <condition attribute=""hil_productcategory"" operator=""eq"" value=""{ProductCategoryGUID}"" />
                    </filter>
                    <link-entity name=""contact"" from=""contactid"" to=""hil_customer"" visible=""false"" link-type=""outer"" alias=""consumer"">
                    <attribute name=""mobilephone"" />
                    </link-entity>
                    </entity>
                    </fetch>";

                    EntityCollection jobsColl = service.RetrieveMultiple(new FetchExpression(fetchCUstomerAsset));
                    if (jobsColl.Entities.Count > 0)
                    {
                        cuatomerAssetColl.Entities.AddRange(jobsColl.Entities);
                        pageNumber++;
                    }
                    else
                    {
                        moreRecord = false;
                    }
                }

                string customerName = "Customer";
                string customerMobileNumber = string.Empty;
                string productsubcategory = string.Empty;
                string productmodel = string.Empty;
                string serialNumber = string.Empty;
                string installation_date = string.Empty;
                string _message = string.Empty;
                EntityCollection smsTempEntCol = null;
                Console.WriteLine("Total Jobs Found " + cuatomerAssetColl.Entities.Count);
                totalCount = cuatomerAssetColl.Entities.Count;
                if (_ModeOfComm == ModeOfCommunication.SMS)
                {
                    QueryExpression _qryExp = new QueryExpression("hil_smstemplates");
                    _qryExp.ColumnSet = new ColumnSet(false);
                    _qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    _qryExp.Criteria.AddCondition("hil_templateid", ConditionOperator.Equal, templateName.Trim());
                    smsTempEntCol = service.RetrieveMultiple(_qryExp);
                }
                logs _logs = new logs();
                D365Campaign d365 = new D365Campaign();
                foreach (Entity job in cuatomerAssetColl.Entities)
                {
                    d365 = new D365Campaign();
                    customerName = "Customer";
                    customerMobileNumber = string.Empty;
                    productsubcategory = string.Empty;
                    productmodel = string.Empty;
                    serialNumber = string.Empty;
                    installation_date = string.Empty;
                    _message = string.Empty;
                    _logs = new logs();
                    try
                    {
                        customerName = job.Contains("hil_customer") ? job.GetAttributeValue<EntityReference>("hil_customer").Name : "";
                        customerMobileNumber = job.GetAttributeValue<AliasedValue>("consumer.mobilephone").Value.ToString();

                        ///////Customer Mobile Number HardCode
                        //customerMobileNumber = "8755904647";

                        productsubcategory = job.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Name;
                        productmodel = job.GetAttributeValue<EntityReference>("msdyn_product").Name;
                        serialNumber = job.GetAttributeValue<String>("msdyn_name").ToString();
                        installation_date = job.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");
                        EntityReference _customer = job.GetAttributeValue<EntityReference>("hil_customer");
                        EntityReference _productDiv = job.GetAttributeValue<EntityReference>("hil_productcategory");
                        _logs.SerialNumber = serialNumber;
                        _logs.MobileNumber = customerMobileNumber;
                        _logs.JobId = "-";
                        _logs.Template = templateName;
                        _logs.ModeOfComunication = _ModeOfComm == ModeOfCommunication.Whatsapp ? "Whats App" : "SMS";
                        if (_ModeOfComm == ModeOfCommunication.Whatsapp)
                        {
                            UserDetails userDetails = new UserDetails();
                            Traits objtraits = new Traits();
                            objtraits.name = customerName;

                            userDetails.phoneNumber = customerMobileNumber;
                            userDetails.countryCode = "+91";
                            userDetails.traits = objtraits;
                            userDetails.tags = new List<object>();

                            UpdateWhatsAppUser(userDetails, _logs);

                            CampaignDetails campaignDetails = new CampaignDetails();
                            campaignDetails.phoneNumber = customerMobileNumber;
                            campaignDetails.countryCode = "+91";
                            campaignDetails.@event = templateName;

                            Traits campaignTraits = new Traits();
                            campaignTraits.expire_on = " ";
                            if (templateName.Contains("_WP_"))// == "5D_WP_AMC"|| templateName == "21D_WP_AMC")
                            {
                                campaignTraits.prd_cat = "Havells Aqua";
                            }
                            if (templateName.Contains("_AC_"))// == "5D_AC_AMC" || templateName == "21D_AC_AMC")
                            {
                                campaignTraits.prd_cat = "Lloyd Air Conditioner";
                            }
                            campaignTraits.product_model = productsubcategory;
                            campaignTraits.product_serial_number = serialNumber;
                            campaignTraits.registration_date = installation_date;

                            campaignDetails.traits = campaignTraits;
                            SendMessage(campaignDetails, _logs);
                            Console.WriteLine("Executing Whatsapp Campaign# " + templateName + " -> " + done.ToString() + "/" + cuatomerAssetColl.Entities.Count + " to " + customerMobileNumber);
                        }
                        else
                        {
                            if (templateName.Trim() == "1107168654225056811")
                            {
                                _message = string.Format("Hi {0},Protect your new Lloyd AC with Havells assured AMC plan. Special price starts at Rs.2499. TnC.  Buy Now https://bit.ly/3py9qIY -Havells", customerName);
                            }
                            else if (templateName.Trim() == "1107169702429550488")
                            {
                                _message = string.Format("Hi {0},Buy Havells AMC plan for your new Havells Water Purifier at flat 20%25 off. TnC. Visit https://bit.ly/3XzawAJ - Havells", customerName);
                            }
                            else
                            {
                                continue;
                            }
                            sendSMS(templateName, customerMobileNumber, _message, _logs);
                            Console.WriteLine("Executing SMS Campaign# " + templateName + " -> " + done.ToString() + "/" + cuatomerAssetColl.Entities.Count + " to " + customerMobileNumber);

                        }
                        done++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("***** " + error + "/" + cuatomerAssetColl.Entities.Count + "Error Occured " + ex.Message);
                        error++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!!!!!!!!!!!!!! Error " + ex.Message);
            }
            Console.WriteLine("===================================================================");
            Console.WriteLine("Summary");
            Console.WriteLine("Template : " + templateName);
            Console.WriteLine("Total Recourd : " + totalCount);
            Console.WriteLine("Skiped : " + skip);
            Console.WriteLine("Send : " + done);
            Console.WriteLine("===================================================================");
            SendEmailOnCampaignCompleation(service, templateName, done, _ModeOfComm == ModeOfCommunication.SMS ? "SMS" : "What's App");
        }
        static protected void sendSMS(string _templateId, string _mobileNumber, string _message, logs _logs)
        {
            string _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=" + _mobileNumber + "&signature=HAVELL&message=" + _message + "&dlt_entity_id=110100001483&dlt_template_id=" + _templateId + "&cust_ref=";

            WebRequest request = WebRequest.Create(_api);
            request.Method = "GET";
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

                msg = _logs.MobileNumber + "," + _logs.SerialNumber + "," + _logs.JobId + "," + _logs.ModeOfComunication + "," + _logs.Template + "," + _api.Replace(',', '`') + "," + "-" + "," + "GATEWAY BUSY APIPayLoad: " + _api;
                CreateOrUpdateLogs(containerName, fileName, msg, header);

            }
            finally
            {
                msg = _logs.MobileNumber + "," + _logs.SerialNumber + "," + _logs.JobId + "," + _logs.ModeOfComunication + "," + _logs.Template + "," + _api.Replace(',', '`') + "," + "-" + "," + responseFromServer;
                CreateOrUpdateLogs(containerName, fileName, msg, header);
            }
        }
        static protected void UpdateWhatsAppUser(UserDetails userDetails, logs _logs)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(service, "Update WhatsApp User");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

            request.AddHeader("Authorization", authInfo);
            request.AddHeader("Content-Type", "application/json");

            request.AddParameter("application/json", data, ParameterType.RequestBody);
            var response = client.Execute(request);

            msg = _logs.MobileNumber + "," + _logs.SerialNumber + "," + _logs.JobId + "," + _logs.ModeOfComunication + "," + _logs.Template + "," + url + "," + data.Replace(',', '`') + "," + response.Content.Replace(',', '`');
            CreateOrUpdateLogs(containerName, fileName, msg, header);
        }
        static protected void SendMessage(CampaignDetails campaignDetails, logs _logs)
        {

            IntegrationConfig intConfig = IntegrationConfiguration(service, "WhatsApp Campaign Details");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(campaignDetails);
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);

            request.AddHeader("Authorization", authInfo);
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("application/json", data, ParameterType.RequestBody);
            var response = client.Execute(request);

            msg = _logs.MobileNumber + "," + _logs.SerialNumber + "," + _logs.JobId + "," + _logs.ModeOfComunication + "," + _logs.Template + "," + url + "," + data.Replace(',', '`') + "," + response.Content.Replace(',', '`');
            CreateOrUpdateLogs(containerName, fileName, msg, header);
        }
        static protected IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + " " + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }

        #region CRM Connection
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

        static private void SendEmailOnCampaignCompleation(IOrganizationService service, string templateName, int totalCount, string ModeOfCommunication)
        {
            EntityCollection entToList = new EntityCollection();
            EntityCollection toTeamMembers = retriveTeamMembers(service, "Consumer Campaign Team");
            foreach (Entity ccEntity in toTeamMembers.Entities)
            {
                Entity entCC = new Entity("activityparty");
                entCC["partyid"] = ccEntity.ToEntityReference();
                entToList.Entities.Add(entCC);
            }
            Entity entFrom = new Entity("activityparty");
            entFrom["partyid"] = new EntityReference("queue", new Guid("9b0480a8-e30f-e911-a94e-000d3af06a98"));
            Entity[] entFromList = { entFrom };

            string mailBody = $@"<div data-wrapper=""true"" dir=""ltr"" style=""font-size:9pt;font-family:'Segoe UI','Helvetica Neue',sans-serif;"">
    <div><span style=""font-family:'Times New Roman',Times,serif""><span style=""font-size:12pt""><span
                    style=""color:maroon"">Dear Team,</span><br><br><span style=""color:maroon"">The {{ModeOfCommunication}}
                    Campaign with template name: {{TemplateName}} is Completed. Please refer to the below Summary:
                </span></span></span><br>&nbsp;<ul>
            <li style=""margin-left: 8px; list-style-position: inside;""><span
                    style=""font-family:'Times New Roman',Times,serif""><span style=""font-size:12pt""><span
                            style=""color:maroon""><strong>Date :</strong>
                            {{date}}</span></span></span></li>
            <li style=""margin-left: 8px; list-style-position: inside;""><span
                    style=""font-family:'Times New Roman',Times,serif""><span style=""font-size:12pt""><span
                            style=""color:maroon""><strong>Total Record Count:</strong>
                            {{TotalRecordCount}}</span></span></span></li>
        </ul><br><span style=""font-family:'Times New Roman',Times,serif""><span style=""font-size:12pt""><span
                    style=""color:maroon"">Please&nbsp;<a href=""{{url}}"">click here</a> to see the log
                    file.</span><br><br><br><span style=""color:maroon"">*** This is system Generated Email, please don’t
                    reply to this mail *** </span><br><br><br><span style=""color:maroon"">Regards,</span><br><span
                    style=""color:maroon"">Team Dynamics</span></span></span></div>
</div>";
            mailBody = mailBody.Replace("{TemplateName}", templateName);
            mailBody = mailBody.Replace("{ModeOfCommunication}", ModeOfCommunication);
            mailBody = mailBody.Replace("{TotalRecordCount}", totalCount.ToString());
            mailBody = mailBody.Replace("{date}", DateTime.Now.ToString("dd-MMMM-yyyy"));
            mailBody = mailBody.Replace("{url}", blobUrl);

            string mailsubject = "The {ModeOfCommunication} Campaign with template name: {TemplateName} on {date} is Completed";


            mailsubject = mailsubject.Replace("{TemplateName}", templateName);
            mailsubject = mailsubject.Replace("{ModeOfCommunication}", ModeOfCommunication);
            mailsubject = mailsubject.Replace("{date}", DateTime.Today.ToString("dd-MMMM-yyyy"));

            Entity email = new Entity("email");
            email["from"] = entFromList;
            email["to"] = entToList;
            email["description"] = mailBody;
            email["subject"] = mailsubject;
            Guid emailId = service.Create(email);

            SendEmailRequest sendEmailReq = new SendEmailRequest()
            {
                EmailId = emailId,
                IssueSend = true
            };
            SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

        }
        static public EntityCollection retriveTeamMembers(IOrganizationService service, string _teamName)
        {
            EntityCollection extTeamMembers = new EntityCollection();
            try
            {
                QueryExpression _query = new QueryExpression("hil_bdteam");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _teamName));
                EntityCollection bdteamCol = service.RetrieveMultiple(_query);
                if (bdteamCol.Entities.Count > 0)
                {
                    QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                    _querymem.ColumnSet = new ColumnSet(false);
                    _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                    _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                    _querymem.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                    EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                    EntityCollection entTOList = new EntityCollection();
                    foreach (Entity entity in bdteammemCol.Entities)
                    {
                        extTeamMembers.Entities.Add(entity);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in Retriving Team Members : " + ex.Message);
            }
            return extTeamMembers;
        }

    }
}
