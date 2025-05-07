using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Webjobs.CX.WAEventForLoyaltyEnrolment
{
    public class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static string _entityName = string.Empty;
        static string _fromDate = string.Empty;
        static string _toDate = string.Empty;
        static string _topRows = string.Empty;
        static string _primaryFieldName = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                PushNotification(_service, "loyalty_enroll_0D"); // Zero Days
                PushNotification5Day(_service, "loyalty_enrollfifteendays"); // Fifteen Days
               // PushNotification5DayBackLog(_service, "loyalty_enrollfifteendays");
                PushNotification7Day(_service, "loyalty_enroll_7D"); // Seven Days
                Console.WriteLine($"Process Datetime:{DateTime.Now.ToString()}");
            }
        }

        static void PushNotification(IOrganizationService _service, string _triggerName)
        {
            try
            {
                if (((CrmServiceClient)_service).IsReady)
                {
                    EntityCollection entcoll = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    DateTime _processDate = DateTime.Now.AddMinutes(330);

                    string _processDateStr = _processDate.ToString("yyyy-MM-dd");
                    string _invoiceDateStart = _processDate.AddDays(-180).ToString("yyyy-MM-dd");
                    string _invoiceDateEnd = _processDate.ToString("yyyy-MM-dd");
                    DateTime _inviteDateCheck = DateTime.Now.Date;
                    string _inviteDateCheckStr = _processDate.ToString("yyyy-MM-dd");
                    string _fetchXML = string.Empty;
                    string _campaignCode = $"WA{_processDateStr.Replace("-", "")}LR0DAY";
                    Console.WriteLine($"Processing 0th Day Product Regustration Whatsapp Invites: {_processDate.ToString()}");
                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='contact'>
                            <attribute name='fullname' />
                            <attribute name='telephone1' />
                            <attribute name='contactid' />
                            <attribute name='mobilephone' />
                            <order attribute='fullname' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_isloyaltyprogramenabled' operator='ne' value='1' />
                              <condition attribute='hil_wacampaignrun' operator='ne' value='{_campaignCode}' /> 
                            </filter>
                            <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='aj'>
                              <attribute name='msdyn_name' />
                              <attribute name='hil_productcategory' /> 
                              <filter type='and'>
                                <filter type='or'>
                                  <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                                  <condition attribute='statuscode' operator='eq' value='910590001' />
                                </filter>
                                <condition attribute='createdon' operator='on' value='{_processDateStr}' />
                                <condition attribute='hil_invoicedate' operator='last-x-days' value='180' />
                              </filter>
                              <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                  <filter type='and'>
                                    <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                  </filter>
                                </link-entity>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
                        Console.WriteLine(_fetchXML);
                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count >= 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    _rowCount++;

                                    string UserName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname") : null;
                                    string MobileNumber = ent.Contains("mobilephone") ? ent.GetAttributeValue<string>("mobilephone") : null;
                                    string Category = ent.Contains("aj.hil_productcategory") ? ((EntityReference)(ent.GetAttributeValue<AliasedValue>("aj.hil_productcategory").Value)).Name : "";
                                    string PrdSerial = ent.Contains("aj.msdyn_name") ? ent.GetAttributeValue < AliasedValue>("aj.msdyn_name").Value.ToString() : "";
                                    var ParamModelData = new TrackUsers();
                                    ParamModelData.phoneNumber = MobileNumber;
                                    ParamModelData.countryCode = "+91";
                                    ParamModelData.traits = new trait();
                                    ParamModelData.traits.name = UserName;

                                    Console.WriteLine($"Processing:{_rowCount}/{_totalCount} Mobile Number:{MobileNumber}");

                                    var RESULT = UpdateWhatsAppUser(ParamModelData);
                                    if (RESULT.result == "true")
                                    {
                                        Console.WriteLine($"Registering Consumer on Whatsapp.");

                                        var ParamData = new TrackEvents();
                                        ParamData.traits = new trait();
                                        ParamData.phoneNumber = MobileNumber;
                                        ParamData.traits.name = UserName;
                                        ParamData.traits.category = Category;
                                        ParamData.traits.prdSerial = PrdSerial;
                                        ParamData.countryCode = "+91";
                                        ParamData.@event = _triggerName;

                                        var resultMsg = SendMessage(ParamData);
                                        if (resultMsg.result == true)
                                        {
                                            Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                                            _entUpdate["hil_wacampaignrun"] = _campaignCode;
                                            _service.Update(_entUpdate);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Whatsapp Invite sent on Mobile Number {MobileNumber} {resultMsg.message}");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Error: While Registering Mobile Number {MobileNumber} {RESULT.message}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Batch Ends.");
                            break;
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        static void PushNotification7Day(IOrganizationService _service, string _triggerName)
        {
            try
            {
                if (((CrmServiceClient)_service).IsReady)
                {
                    EntityCollection entcoll = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    DateTime _processDateNow = DateTime.Now.AddMinutes(330);
                    DateTime _processDate = DateTime.Now.AddMinutes(330).AddDays(-7);//new DateTime(2024, 9, 2).AddDays(-15);
                    string _processDateStr = _processDate.ToString("yyyy-MM-dd");
                    string _processDateNowStr = _processDateNow.ToString("yyyy-MM-dd");
                    string _invoiceDateStart = _processDate.AddDays(-180).ToString("yyyy-MM-dd");
                    string _invoiceDateEnd = _processDate.ToString("yyyy-MM-dd");
                    DateTime _inviteDateCheck = DateTime.Now.Date;
                    string _inviteDateCheckStr = _processDate.ToString("yyyy-MM-dd");
                    string _fetchXML = string.Empty;
                    string _campaignCode = $"WA{_processDateNowStr.Replace("-", "")}LR7DAY";
                    Console.WriteLine("Processing 7th Day Product Registration SMS Invites.");
                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                        <entity name='contact'>
                                        <attribute name='fullname' />
                                        <attribute name='telephone1' />
                                        <attribute name='contactid' />
                                        <attribute name='mobilephone' />
                                        <order attribute='fullname' descending='false' />
                                        <filter type='and'>
                                        <condition attribute='hil_isloyaltyprogramenabled' operator='ne' value='1' />
                                        <condition attribute='hil_wacampaignrun2' operator='ne' value='{_campaignCode}' />
                                        </filter>
                                        <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='aj'>
                                        <filter type='and'>
                                        <filter type='or'>
                                        <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                                        <condition attribute='statuscode' operator='eq' value='910590001' />
                                        </filter>
                                        <condition attribute='createdon' operator='on' value='{_processDateStr}' />
                                        <condition attribute='hil_invoicedate' operator='on-or-after' value='{_invoiceDateStart}' />
                                        <condition attribute='hil_invoicedate' operator='on-or-before' value='{_invoiceDateEnd}' />
                                        </filter>
                                        <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                                        <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                        <filter type='and'>
                                        <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                        </filter>
                                        </link-entity>
                                        </link-entity>
                                        </link-entity>
                                        </entity>
                                        </fetch>";
                        Console.WriteLine("***************Query Starts***********");
                        Console.WriteLine(_fetchXML);
                        Console.WriteLine("***************Query Ends*************");
                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count >= 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    _rowCount++;
                                    string UserName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname") : null;
                                    string MobileNumber = ent.Contains("mobilephone") ? ent.GetAttributeValue<string>("mobilephone") : null;

                                    Console.WriteLine($"Processing:{_rowCount}/{_totalCount} Mobile Number:{MobileNumber}");

                                    Entity smsEntity = new Entity("hil_smsconfiguration");
                                    smsEntity["hil_smstemplate"] = new EntityReference("hil_smstemplates", new Guid("03d202cc-c2b2-ef11-b8e8-0022486eb03e"));
                                    smsEntity["subject"] = $"Loyalty Trigger 7 days";
                                    smsEntity["hil_message"] = $"Join Havells Happiness Loyalty Program today!Earn sign up bonus, points on every purchase %26 unlock exclusive offers.Sign Up now https://wa.me/9711773333%3ftext=Join T&C apply -Havells";
                                    smsEntity["hil_mobilenumber"] = MobileNumber;
                                    smsEntity["hil_direction"] = new OptionSetValue(2); // Outgoing                   
                                    _service.Create(smsEntity);

                                    Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                                    _entUpdate["hil_wacampaignrun2"] = _campaignCode;
                                    _service.Update(_entUpdate);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Batch Ends.");
                            break;
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        static void PushNotificationBackup(IOrganizationService _service,string _triggerName)
        {
            try
            {
                if (((CrmServiceClient)_service).IsReady)
                {
                    EntityCollection entcoll = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    DateTime _processDate = new DateTime(2024, 8, 28);
                    string _processDateStr = _processDate.ToString("yyyy-MM-dd");
                    string _invoiceDateStart  = _processDate.AddDays(-180).ToString("yyyy-MM-dd");
                    string _invoiceDateEnd = _processDate.ToString("yyyy-MM-dd");
                    DateTime _inviteDateCheck = DateTime.Now.Date;
                    string _inviteDateCheckStr = _processDate.ToString("yyyy-MM-dd");
                    string _fetchXML = string.Empty;
                    Console.WriteLine("Processing 0th Day Product Regustration Whatsapp Invites.");
                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='contact'>
                            <attribute name='fullname' />
                            <attribute name='telephone1' />
                            <attribute name='contactid' />
                            <attribute name='mobilephone' />
                            <order attribute='fullname' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_loyaltyprogramenabled' operator='ne' value='1' />
                              <condition attribute='hil_wacampaignrun' operator='ne' value='WA20240828LR0DAY' />
                            </filter>
                            <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='aj'>
                              <filter type='and'>
                                <filter type='or'>
                                  <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                                  <condition attribute='statuscode' operator='eq' value='910590001' />
                                </filter>
                                <condition attribute='createdon' operator='on' value='{_processDateStr}' />
                                <condition attribute='hil_invoicedate' operator='on-or-after' value='{_invoiceDateStart}' />
                                <condition attribute='hil_invoicedate' operator='on-or-before' value='{_invoiceDateEnd}' />
                              </filter>
                              <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                  <filter type='and'>
                                    <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                  </filter>
                                </link-entity>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count >= 0) {
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    _rowCount++;
                                    
                                    string UserName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname") : null;
                                    string MobileNumber = ent.Contains("mobilephone") ? ent.GetAttributeValue<string>("mobilephone") : null;
                                    var ParamModelData = new TrackUsers();
                                    ParamModelData.phoneNumber = MobileNumber;
                                    ParamModelData.countryCode = "+91";
                                    ParamModelData.traits = new trait();
                                    ParamModelData.traits.name = UserName;

                                    Console.WriteLine($"Processing:{_rowCount}/{_totalCount} Mobile Number:{MobileNumber}");

                                    var RESULT = UpdateWhatsAppUser(ParamModelData);
                                    if (RESULT.result == "true")
                                    {
                                        Console.WriteLine($"Registering Consumer on Whatsapp.");

                                        var ParamData = new TrackEvents();
                                        ParamData.traits = new trait();
                                        ParamData.phoneNumber = MobileNumber;
                                        ParamData.traits.name = UserName;
                                        ParamData.countryCode = "+91";
                                        ParamData.@event = _triggerName;

                                        var resultMsg = SendMessage(ParamData);
                                        if (resultMsg.result == true)
                                        {
                                            Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                                            _entUpdate["hil_wacampaignrun"] = "WA20240828LR0DAY";
                                            _service.Update(_entUpdate);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Whatsapp Invite sent on Mobile Number {MobileNumber} {resultMsg.message}");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Error: While Registering Mobile Number {MobileNumber} {RESULT.message}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                        else {
                            Console.WriteLine("Batch Ends.");
                            break;
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        static void PushNotification5Day(IOrganizationService _service, string _triggerName)
        {
            try
            {
                if (((CrmServiceClient)_service).IsReady)
                {
                    EntityCollection entcoll = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    DateTime _processDateNow = DateTime.Now.AddMinutes(330);
                    DateTime _processDate = DateTime.Now.AddMinutes(330).AddDays(-5);//new DateTime(2024, 9, 2).AddDays(-15);
                    string _processDateStr = _processDate.ToString("yyyy-MM-dd");
                    string _processDateNowStr = _processDateNow.ToString("yyyy-MM-dd");
                    string _invoiceDateStart = _processDate.AddDays(-180).ToString("yyyy-MM-dd");
                    string _invoiceDateEnd = _processDate.ToString("yyyy-MM-dd");
                    DateTime _inviteDateCheck = DateTime.Now.Date;
                    string _inviteDateCheckStr = _processDate.ToString("yyyy-MM-dd");
                    string _fetchXML = string.Empty;
                    string _campaignCode = $"WA{_processDateNowStr.Replace("-", "")}LR5DAY";
                    Console.WriteLine("Processing 15th Day Product Regustration Whatsapp Invites.");
                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='contact'>
                            <attribute name='fullname' />
                            <attribute name='telephone1' />
                            <attribute name='contactid' />
                            <attribute name='mobilephone' />
                            <order attribute='fullname' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_isloyaltyprogramenabled' operator='ne' value='1' />
                              <condition attribute='hil_wacampaignrun2' operator='ne' value='{_campaignCode}' />
                            </filter>
                            <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='aj'>
                              <filter type='and'>
                                <filter type='or'>
                                  <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                                  <condition attribute='statuscode' operator='eq' value='910590001' />
                                </filter>
                                <condition attribute='createdon' operator='on' value='{_processDateStr}' />
                                <condition attribute='hil_invoicedate' operator='on-or-after' value='{_invoiceDateStart}' />
                                <condition attribute='hil_invoicedate' operator='on-or-before' value='{_invoiceDateEnd}' />
                              </filter>
                              <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                  <filter type='and'>
                                    <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                  </filter>
                                </link-entity>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
                        Console.WriteLine("***************Query Starts***********");
                        Console.WriteLine(_fetchXML);
                        Console.WriteLine("***************Query Ends*************");
                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count >= 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    _rowCount++;

                                    string UserName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname") : null;
                                    string MobileNumber = ent.Contains("mobilephone") ? ent.GetAttributeValue<string>("mobilephone") : null;
                                    var ParamModelData = new TrackUsers();
                                    ParamModelData.phoneNumber = MobileNumber;
                                    ParamModelData.countryCode = "+91";
                                    ParamModelData.traits = new trait();
                                    ParamModelData.traits.name = UserName;

                                    Console.WriteLine($"Processing:{_rowCount}/{_totalCount} Mobile Number:{MobileNumber}");

                                    var RESULT = UpdateWhatsAppUser(ParamModelData);
                                    if (RESULT.result == "true")
                                    {
                                        Console.WriteLine($"Registering Consumer on Whatsapp.");

                                        var ParamData = new TrackEvents();
                                        ParamData.traits = new trait();
                                        ParamData.phoneNumber = MobileNumber;
                                        ParamData.traits.name = UserName;
                                        ParamData.countryCode = "+91";
                                        ParamData.@event = _triggerName;

                                        var resultMsg = SendMessage(ParamData);
                                        if (resultMsg.result == true)
                                        {
                                            Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                                            _entUpdate["hil_wacampaignrun2"] = _campaignCode;
                                            _service.Update(_entUpdate);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Whatsapp Invite sent on Mobile Number {MobileNumber} {resultMsg.message}");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Error: While Registering Mobile Number {MobileNumber} {RESULT.message}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Batch Ends.");
                            break;
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        static void PushNotification5DayBackLog(IOrganizationService _service, string _triggerName)
        {
            try
            {
                if (((CrmServiceClient)_service).IsReady)
                {
                    EntityCollection entcoll = null;
                    int _rowCount = 0;
                    int _totalCount = 0;
                    string _datetime = ConfigurationManager.AppSettings["powerapp:processDate"].ToString();
                    DateTime _processDateNow = Convert.ToDateTime(_datetime).AddMinutes(330);
                    DateTime _processDate = _processDateNow.AddDays(-5);
                    string _processDateStr = _processDate.ToString("yyyy-MM-dd");
                    string _processDateNowStr = _processDateNow.ToString("yyyy-MM-dd");
                    string _invoiceDateStart = _processDate.AddDays(-180).ToString("yyyy-MM-dd");
                    string _invoiceDateEnd = _processDate.ToString("yyyy-MM-dd");
                    DateTime _inviteDateCheck = DateTime.Now.Date;
                    string _inviteDateCheckStr = _processDate.ToString("yyyy-MM-dd");
                    string _fetchXML = string.Empty;
                    string _campaignCode = $"WA{_processDateNowStr.Replace("-", "")}LR5DAY";
                    Console.WriteLine("Processing 5th Day Product Regustration Whatsapp Invites.");
                    do
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='contact'>
                            <attribute name='fullname' />
                            <attribute name='telephone1' />
                            <attribute name='contactid' />
                            <attribute name='mobilephone' />
                            <order attribute='fullname' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_loyaltyprogramenabled' operator='ne' value='1' />
                              <condition attribute='hil_wacampaignrun2' operator='ne' value='{_campaignCode}' /> 
                            </filter>
                            <link-entity name='msdyn_customerasset' from='hil_customer' to='contactid' link-type='inner' alias='aj'>
                              <filter type='and'>
                                <filter type='or'>
                                  <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                                  <condition attribute='statuscode' operator='eq' value='910590001' />
                                </filter>
                                <condition attribute='createdon' operator='on' value='{_processDateStr}' />
                                <condition attribute='hil_invoicedate' operator='on-or-after' value='{_invoiceDateStart}' />
                                <condition attribute='hil_invoicedate' operator='on-or-before' value='{_invoiceDateEnd}' />
                              </filter>
                              <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                  <filter type='and'>
                                    <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                  </filter>
                                </link-entity>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
                        Console.WriteLine("***************Query Starts***********");
                        Console.WriteLine(_fetchXML);
                        Console.WriteLine("***************Query Ends*************");
                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count >= 0)
                        {
                            _totalCount += entcoll.Entities.Count;
                            foreach (Entity ent in entcoll.Entities)
                            {
                                try
                                {
                                    _rowCount++;

                                    string UserName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname") : null;
                                    string MobileNumber = ent.Contains("mobilephone") ? ent.GetAttributeValue<string>("mobilephone") : null;
                                    var ParamModelData = new TrackUsers();
                                    ParamModelData.phoneNumber = MobileNumber;
                                    ParamModelData.countryCode = "+91";
                                    ParamModelData.traits = new trait();
                                    ParamModelData.traits.name = UserName;

                                    Console.WriteLine($"Processing:{_rowCount}/{_totalCount} Mobile Number:{MobileNumber}");

                                    var RESULT = UpdateWhatsAppUser(ParamModelData);
                                    if (RESULT.result == "true")
                                    {
                                        Console.WriteLine($"Registering Consumer on Whatsapp.");

                                        var ParamData = new TrackEvents();
                                        ParamData.traits = new trait();
                                        ParamData.phoneNumber = MobileNumber;
                                        ParamData.traits.name = UserName;
                                        ParamData.countryCode = "+91";
                                        ParamData.@event = _triggerName;

                                        var resultMsg = SendMessage(ParamData);
                                        if (resultMsg.result == true)
                                        {
                                            Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                                            _entUpdate["hil_wacampaignrun2"] = _campaignCode;
                                            _service.Update(_entUpdate);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Error: While Whatsapp Invite sent on Mobile Number {MobileNumber} {resultMsg.message}");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Error: While Registering Mobile Number {MobileNumber} {RESULT.message}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Batch Ends.");
                            break;
                        }
                    }
                    while (entcoll.MoreRecords);
                    Console.WriteLine("Batch Ends.");
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        static sendMesgRes SendMessage(TrackEvents campaignDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "WhatsApp Campaign Details");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string Campaigndata = JsonConvert.SerializeObject(campaignDetails);
            using (var client = new HttpClient())
            {
                HttpContent objCampaign = new StringContent(Campaigndata);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objCampaign).Result;
                return JsonConvert.DeserializeObject<sendMesgRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        static UpdateAppUserRes UpdateWhatsAppUser(TrackUsers userDetails)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(_service, "Update WhatsApp User");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            using (var client = new HttpClient())
            {
                HttpContent objcontent = new StringContent(data);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objcontent).Result;
                return JsonConvert.DeserializeObject<UpdateAppUserRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
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
    }
    public class TrackUsers
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string tags { get; set; }
        public trait traits { get; set; }
    }
    public class TrackEvents
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string @event { get; set; }
        public trait traits { get; set; }
    }
    public class trait
    {
        public string name { get; set; }
        public string category { get; set; }
        public string prdSerial { get; set; }
    }
    public class UpdateAppUserRes
    {
        public string result { get; set; }
        public string message { get; set; }
    }
    public class sendMesgRes
    {
        public bool result { get; set; }
        public string message { get; set; }
        public string id { get; set; }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
}
