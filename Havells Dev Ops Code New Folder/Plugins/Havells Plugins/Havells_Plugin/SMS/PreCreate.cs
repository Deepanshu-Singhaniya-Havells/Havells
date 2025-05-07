using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;
using Havells_Plugin;
using System.Net;

namespace Havells_Plugin.SMS
{
    public class PreCreate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_smsconfiguration"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity _smsTemplate = null;
                    Guid _regardingObjectId = Guid.Empty;

                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    //Check the direction if it's incoming
                    OptionSetValue Direction = null;
                    Direction= entity.GetAttributeValue<OptionSetValue>("hil_direction");
                    
                    if (Direction != null && Direction.Value == 1)//Incoming
                    {
                        tracingService.Trace("2");
                        DecodeCustomerandMessage(entity, service);
                    }
                    if (Direction != null && Direction.Value == 2)//OutGoing
                    {
                        tracingService.Trace("3");
                        string _smsTemplateId = string.Empty;

                        if (entity.Attributes.Contains("hil_smstemplate")) {
                            EntityReference _entSMSTemplate = entity.GetAttributeValue<EntityReference>("hil_smstemplate");
                            _smsTemplate = service.Retrieve("hil_smstemplates", _entSMSTemplate.Id, new ColumnSet("hil_templateid", "hil_encryptsms"));
                            _smsTemplateId = _smsTemplate.GetAttributeValue<string>("hil_templateid");
                        }
                        if (entity.Attributes.Contains("regardingobjectid"))
                        {
                            _regardingObjectId = entity.GetAttributeValue<EntityReference>("regardingobjectid").Id;
                        }

                        //_regardingObjectId
                        hil_smsconfiguration SMS = entity.ToEntity<hil_smsconfiguration>();
                        
                        if (SMS.hil_MobileNumber != null && SMS.hil_Message != null)
                        {
                            string kkgOTP = string.Empty;
                            string kkgOTPOriginal = string.Empty;
                            string message = SMS.hil_Message;
                            string _temp = "";
                            string _custref = string.Empty;

                            if (SMS.hil_requesttype == "KKGCode" || message.IndexOf("<KKGCode>") >0)
                            {
                                //kkgOTPOriginal = GetKKGOTP(service, SMS.RegardingObjectId.Id);
                                //kkgOTP = Havells_Plugin.WorkOrder.Common.Base64Decode(kkgOTPOriginal);
                                //if (kkgOTP.Trim().Length == 4) //For Descrypted KKG Codes
                                //{
                                //    message = message.Replace("<KKGCode>", kkgOTP);
                                //} 
                                //else { //For Not Descrypted KKG Codes
                                //    message = message.Replace("<KKGCode>", kkgOTPOriginal);
                                //}
                                string _KKGCode = KKGCodeHashing.GenerateKKGCodeHash(SMS.RegardingObjectId.Name, service);
                                tracingService.Trace("KKG Code Generation... " + _KKGCode);
                                message = message.Replace("<KKGCode>", _KKGCode);
                            }
                            else if (SMS.hil_requesttype == "NPS")
                            {
                                _custref = SMS.ActivityAdditionalParams;
                            }
                            message = message.Replace("#", "%23");
                            message = message.Replace("&", "%26");
                            message = message.Replace("+", "%2B");
                            if (_smsTemplateId != string.Empty)
                            {
                                bool _duplicateDetectionRule = false;

                                Entity entTemp = service.Retrieve("hil_smstemplates", entity.GetAttributeValue<EntityReference>("hil_smstemplate").Id, new ColumnSet("hil_enableduplicatedetectionrule"));
                                if (entTemp != null)
                                {
                                    _duplicateDetectionRule = entTemp.GetAttributeValue<bool>("hil_enableduplicatedetectionrule");
                                }
                                if (_duplicateDetectionRule)
                                {
                                    QueryExpression query = new QueryExpression(entity.LogicalName)
                                    {
                                        ColumnSet = new ColumnSet(false),
                                        Criteria = new FilterExpression(LogicalOperator.And)
                                    };
                                    query.Criteria.AddCondition("createdon", ConditionOperator.LastXHours, 1); // 60 Minutes
                                    query.Criteria.AddCondition("hil_smstemplate", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("hil_smstemplate").Id);
                                    query.Criteria.AddCondition("hil_mobilenumber", ConditionOperator.Equal, SMS.hil_MobileNumber);
                                    query.Criteria.AddCondition("hil_direction", ConditionOperator.Equal, 2);
                                    if (entity.Contains("regardingobjectid"))
                                    {
                                        query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("regardingobjectid").Id);
                                    }
                                    EntityCollection entSMS = service.RetrieveMultiple(query);

                                    if (entSMS.Entities.Count == 0)
                                    {
                                        if (_smsTemplateId == "1107161191448698079" || _smsTemplateId == "1107161191438154934") //KKG Code Send and Resend Templates
                                            SMS = HelperShootSMS.OnDemandSMSShootFunctionDLT(service, message, SMS.hil_MobileNumber, SMS, _smsTemplateId, _custref, _regardingObjectId);
                                    }
                                }
                                else
                                {
                                    if (_smsTemplateId == "1107161191448698079" || _smsTemplateId == "1107161191438154934") //KKG Code Send and Resend Templates
                                        SMS = HelperShootSMS.OnDemandSMSShootFunctionDLT(service, message, SMS.hil_MobileNumber, SMS, _smsTemplateId, _custref, _regardingObjectId);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SMS_PreOp_Create.ResolveMessage" + ex.Message);
            }
        }
        public static string GetKKGOTP(IOrganizationService service, Guid workOrderId)
        {
            string _kkgOPT = string.Empty;
            try
            {
                QueryExpression Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("hil_kkgotp");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_workorderid", ConditionOperator.Equal, workOrderId));
                EntityCollection enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    _kkgOPT = enCol.Entities[0].GetAttributeValue<string>("hil_kkgotp");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SMS_PreOp_Create.Execute.GetKKGOTP" + ex.Message);
            }
            return _kkgOPT;
        }
        #region REGEX
        public static void DecodeCustomerandMessage(Entity SMS, IOrganizationService service)
        {
            Boolean isFormatCorrect = false;
            OptionSetValue CreateRecord = null;
            String txt = String.Empty;
            String mobileNumber = String.Empty;
            String sMobileNo = String.Empty;
            Guid fsAccountId = Guid.Empty;
            Guid fsContactId = Guid.Empty;
            Guid fsPinCodeId = Guid.Empty;
            try
            {
                if (SMS.Contains("hil_message"))
                {
                    txt = SMS.GetAttributeValue<String>("hil_message");
                }
                if (SMS.Contains("hil_mobilenumber"))
                {
                    mobileNumber = SMS.GetAttributeValue<String>("hil_mobilenumber");
                }
                txt = txt.Replace(" ", "");
                txt = txt.ToUpper();

                if (txt == "SCORE") // SCORE
                {
                    long JobIdVer = 0;
                    int iKKGVer = 0;
                    string KKG = txt.Substring(txt.Length - 4);
                    string JobId = txt.Substring(1, txt.Length - 5);
                    bool result = long.TryParse(JobId, out JobIdVer);
                    if (result)
                    {
                        result = int.TryParse(KKG, out iKKGVer);
                        if (result)
                        {
                            UpdateWorkOrderSMS(service, SMS, KKG, JobId);
                        }
                    }
                }
                else if (txt.StartsWith("C")) // Work Done SMS
                {
                    tracingService.Trace("3");
                    //string[] sms = txt.Split(' ');
                    long JobIdVer = 0;
                    int iKKGVer = 0;
                    string KKG = txt.Substring(txt.Length - 6);
                    string JobId = txt.Substring(1, txt.Length - 7);
                    tracingService.Trace("4");
                    bool result = long.TryParse(JobId, out JobIdVer);
                    if (result)
                    {
                        result = int.TryParse(KKG, out iKKGVer);
                        if (result)
                        {
                            tracingService.Trace("7");
                            UpdateWorkOrderSMS(service, SMS, KKG, JobId);
                            tracingService.Trace("8");
                        }
                    }
                }
                else if (txt.StartsWith("R")) // Resend KKG Code
                {
                    long JobIdVer = 0;
                    string opCode = txt.Substring(0, 4);
                    if (opCode == "RKKG")
                    {
                        string JobId = txt.Substring(4, txt.Length - 4);
                        bool result = long.TryParse(JobId, out JobIdVer);
                        if (result)
                        {
                            QueryExpression Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet(false);
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, JobId));
                            EntityCollection enCol = service.RetrieveMultiple(Query);
                            if (enCol.Entities.Count > 0)
                            {
                                msdyn_workorder enJob = enCol.Entities[0].ToEntity<msdyn_workorder>();
                                enJob["hil_resendkkg"] = true;
                                service.Update(enJob);
                                SMS["regardingobjectid"] = new EntityReference(msdyn_workorder.EntityLogicalName, enJob.Id);
                            }
                        }
                    }
                }
                else
                {
                    #region ResolveCustomer
                    if (SMS.Contains("hil_mobilenumber"))
                    {
                        sMobileNo = SMS.GetAttributeValue<String>("hil_mobilenumber");
                        sMobileNo = sMobileNo.Trim().Length > 10 ? sMobileNo.Trim().Substring(2, 10) : sMobileNo.Trim();
                    }
                    fsAccountId = Helper.GetGuidbyName("account", "telephone1", sMobileNo, service);
                    fsContactId = Helper.GetGuidbyName("contact", "mobilephone", sMobileNo, service);
                    if (fsContactId == Guid.Empty)
                    {
                        //Create Contact
                        Entity Contact = new Entity("contact");
                        Contact["firstname"] = "tobe";
                        Contact["lastname"] = "updated";
                        Contact["mobilephone"] = sMobileNo;
                        fsContactId = service.Create(Contact);
                    }

                    if (fsAccountId != Guid.Empty)
                        SMS["hil_account"] = new EntityReference("account", fsAccountId);
                    if (fsContactId != Guid.Empty)
                        SMS["hil_contact"] = new EntityReference("contact", fsContactId);
                    #endregion
                    txt = txt.Substring(0, 6);
                    if (txt == string.Empty) return;
                    #region MatchRegex
                    string re2 = "(\\d)";   // Any Single Digit 1
                    string re3 = "(\\d)";   // Any Single Digit 2
                    string re4 = "(\\d)";   // Any Single Digit 3
                    string re5 = "(\\d)";   // Any Single Digit 4
                    string re6 = "(\\d)";   // Any Single Digit 5
                    string re7 = "(\\d)";   // Any Single Digit 6
                    String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_integrationconfiguration'>
                    <attribute name='hil_integrationconfigurationid' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_requesttype' />
                    <attribute name='hil_createrecord' />
                    <attribute name='hil_brandidentifier' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_name' operator='eq' value='SMS_Integration_Incoming_RegEx' />
                    </filter>
                  </entity>
                </fetch>";
                    EntityCollection enColl = service.RetrieveMultiple(new FetchExpression(fetch));
                    foreach (Entity en in enColl.Entities)
                    {
                        String sBrandIdentifier = String.Empty;
                        String sReqType = String.Empty;
                        if (en.Contains("hil_brandidentifier"))
                        {
                            sBrandIdentifier = en.GetAttributeValue<String>("hil_brandidentifier");
                            sBrandIdentifier = sBrandIdentifier.Replace(" ", "");
                        }
                        if (en.Contains("hil_requesttype"))
                        {
                            sReqType = en.GetAttributeValue<String>("hil_requesttype");
                            sReqType = sReqType.Replace(" ", "");
                        }
                        string re1 = "(" + sBrandIdentifier + ")"; // Any Single Word Character (Not Whitespace) 1
                        string re8 = "(" + sReqType + ")";  // Word 1
                        Regex r = new Regex(re2 + re3 + re4 + re5 + re6 + re7, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        Match m = r.Match(txt);
                        if (m.Success && (txt.Length == 6))
                        {
                            if (en.Contains("hil_createrecord"))
                            {
                                CreateRecord = en.GetAttributeValue<OptionSetValue>("hil_createrecord");
                            }
                            isFormatCorrect = true;
                            break;
                        }
                    }
                    #endregion
                    if (isFormatCorrect)
                    {
                        if (CreateRecord != null)
                        {
                            SMS["hil_createrecord"] = CreateRecord;//Create record type i.e. Case/Lead
                        }
                        //H250343
                        //string sBrandIdentifier = txt.Substring(0, 1);
                        string sPinCode = txt;
                        //string sRequestType = txt.Substring(1, 6);
                        //SMS["hil_brandidentifier"] = sBrandIdentifier;
                        SMS["hil_pincode"] = sPinCode;
                        //SMS["hil_requesttype"] = sRequestType;
                        SMS["statuscode"] = new OptionSetValue(910590003);//Incoming SMS format matched
                        fsPinCodeId = Helper.GetGuidbyName("hil_pincode", "hil_name", sPinCode, service);
                        if (fsPinCodeId != Guid.Empty)
                            SMS["hil_pincoderef"] = new EntityReference("hil_pincode", fsPinCodeId);
                        string ReqType = GetRequestType(service, SMS);
                        if (ReqType == "JOB")
                        {
                            SMS["hil_requesttype"] = ReqType;
                            msdyn_workorder iJob = new msdyn_workorder();
                            iJob.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, fsPinCodeId);
                            iJob.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, fsContactId);
                            if (SMS.Contains("hil_mobilenumber"))
                            {
                                sMobileNo = SMS.GetAttributeValue<String>("hil_mobilenumber");
                                sMobileNo = sMobileNo.Trim().Length > 10 ? sMobileNo.Trim().Substring(2, 10) : sMobileNo.Trim();
                                iJob.hil_mobilenumber = sMobileNo;
                            }
                            Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
                            Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                            iJob.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                            iJob.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                            iJob.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
                            iJob.hil_quantity = Convert.ToInt32(1);
                            iJob.hil_SourceofJob = new OptionSetValue(5);
                            if (SMS.Contains("hil_message"))
                            {
                                txt = SMS.GetAttributeValue<String>("hil_message");
                                iJob.hil_CustomerComplaintDescription = txt;
                            }
                            iJob.hil_Brand = new OptionSetValue(1); //Havells
                            //iJob.hil_AutomaticAssign = new OptionSetValue(2);
                            Guid WorkOd = service.Create(iJob);
                            SMS["regardingobjectid"] = new EntityReference(msdyn_workorder.EntityLogicalName, WorkOd);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", WorkOd, service);
                        }
                        else
                        {
                            SMS["hil_requesttype"] = ReqType;
                        }
                        //service.Update(SMS);
                    }
                    else
                    {
                        SMS["statuscode"] = new OptionSetValue(910590002); //Incoming SMS format didn't matched
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SMS_PreOp_Create.ResolveMessage" + ex.Message);
            }
        }
        public static string GetRequestType(IOrganizationService service, Entity SMS)
        {
            string TOMob = string.Empty;
            string ReqType = string.Empty;
            if (SMS.Contains("hil_tomobile") && SMS.Attributes.Contains("hil_tomobile"))
            {
                TOMob = SMS.GetAttributeValue<string>("hil_tomobile");
                if (TOMob != null)
                {
                    QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("hil_requesttype");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_url", ConditionOperator.Equal, TOMob));
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        hil_integrationconfiguration IConf = Found.Entities[0].ToEntity<hil_integrationconfiguration>();
                        if (IConf.hil_requesttype != null)
                        {
                            ReqType = IConf.hil_requesttype;
                        }
                    }
                }
            }
            return ReqType;
        }
        #endregion
        public static void UpdateWorkOrderSMS(IOrganizationService service, Entity SMS, string KKG, string JobNo)
        {
            msdyn_workorder enUpdateJob = new msdyn_workorder();
            EntityReference reSubStatus = new EntityReference();
            string iKKG = string.Empty;
            QueryExpression Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("msdyn_substatus");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, JobNo));
            EntityCollection enCol = service.RetrieveMultiple(Query);
            if(enCol.Entities.Count > 0)
            {
                msdyn_workorder enJob = enCol.Entities[0].ToEntity<msdyn_workorder>();
                if(enJob.msdyn_SubStatus != null)
                {
                    reSubStatus = enJob.msdyn_SubStatus;
                    if (reSubStatus.Name != "Closed" && reSubStatus.Name != "Work Done" && reSubStatus.Name != "Work Done SMS")
                    {

                        //iKKG = enJob.hil_KKGOTP;
                        //Added by Kuldeep Khare on 25/Dec/2019 KKG Code encryption
                        //if (iKKG != null && (iKKG == KKG)
                        //if (iKKG != null && (iKKG == KKG || Havells_Plugin.WorkOrder.Common.Base64Decode(iKKG) == KKG))
                        //{
                        enUpdateJob.Id = enJob.Id;
                        //enUpdateJob.hil_kkgcode = iKKG;
                        enUpdateJob.hil_kkgcode = KKG;
                        enUpdateJob["hil_tattime"] = DateTime.Now.AddMinutes(330);
                        DateTime CreatedOn = new DateTime();
                        DateTime WorkDoneOnSMS = new DateTime();
                        CreatedOn = Convert.ToDateTime(enJob.CreatedOn);
                        WorkDoneOnSMS = Convert.ToDateTime(enUpdateJob["hil_tattime"]);
                        TimeSpan diff = WorkDoneOnSMS - CreatedOn;
                        double hours = diff.TotalMinutes / 60;
                        enUpdateJob["hil_tattimecalculated"] = Convert.ToDecimal(hours);
                        EntityReference entRef = Havells_Plugin.WorkOrder.Common.TATCategory(service, hours);
                        if (entRef != null)
                        {
                            enUpdateJob["hil_tatcategory"] = entRef;
                        }
                        enUpdateJob["hil_requesttype"] = new OptionSetValue(1);
                        service.Update(enUpdateJob);
                        CalculateChargesINSMS(service, enJob.Id);
                        Havells_Plugin.WorkOrder.PostUpdate.CalculateTAT(service, enJob.Id);
                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Done SMS", enJob.Id, service);
                        SMS["regardingobjectid"] = new EntityReference(msdyn_workorder.EntityLogicalName, enJob.Id);

                        //service.Update(enUpdateJob);
                        //}
                    }
                }
            }
        }
        #region Calculate Charges
        public static void CalculateChargesINSMS(IOrganizationService service, Guid TicketId)
        {
            decimal TotalCharges = 0;
            msdyn_workorder Ticket = new msdyn_workorder();//new msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, TicketId, new ColumnSet(false));
            Ticket.Id = TicketId;
            QueryExpression Query = new QueryExpression(hil_estimate.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_totalcharges");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_job", ConditionOperator.Equal, TicketId);
            EntityCollection Found1 = service.RetrieveMultiple(Query);
            if (Found1.Entities.Count > 0)
            {
                hil_estimate iEstimate = Found1.Entities[0].ToEntity<hil_estimate>();
                if (iEstimate.hil_totalCharges != null)
                {
                    Ticket.hil_actualcharges = new Money(iEstimate.hil_totalCharges.Value);
                    Ticket["hil_payblechargedecimal"] = (decimal)iEstimate.hil_totalCharges;
                    Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                    //Ticket.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330); 
                    if (iEstimate.hil_totalCharges > 0)
                        Ticket["hil_ischargable"] = true;
                    else
                        Ticket["hil_ischargable"] = false;
                    //Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                    service.Update(Ticket);
                }
                else
                {
                    if (iEstimate.hil_ServiceCharges != null)
                    {
                        TotalCharges = TotalCharges + iEstimate.hil_ServiceCharges.Value;
                    }
                    if (iEstimate.hil_PartCharges != null)
                    {
                        TotalCharges = TotalCharges + iEstimate.hil_PartCharges.Value;
                    }
                    Ticket.hil_actualcharges = new Money(TotalCharges);
                    Ticket["hil_payblechargedecimal"] = (decimal)TotalCharges;
                    if (TotalCharges > 0)
                        Ticket["hil_ischargable"] = true;
                    else
                        Ticket["hil_ischargable"] = false;
                    //Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                    //Ticket.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330);
                    service.Update(Ticket);
                }
            }
            else
            {
                QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderincident.EntityLogicalName);
                Qry.ColumnSet = new ColumnSet("hil_warrantystatus", "msdyn_customerasset", "statecode");
                Qry.AddAttributeValue("msdyn_workorder", TicketId);
                EntityCollection Found = service.RetrieveMultiple(Qry);
                if (Found.Entities.Count >= 1)
                {
                    foreach (msdyn_workorderincident Inc in Found.Entities)
                    {
                        if (Inc.statecode.Equals(msdyn_workorderincidentState.Active))
                        {
                            Money TempCharges = CalculateActualChargesOutWarranty(service, Inc.Id, TicketId); ;
                            TotalCharges = TotalCharges + TempCharges.Value;
                        }
                    }
                }
                Ticket.hil_actualcharges = new Money(TotalCharges);
                Ticket["hil_payblechargedecimal"] = (decimal)TotalCharges;
                if (TotalCharges > 0)
                    Ticket["hil_ischargable"] = true;
                else
                    Ticket["hil_ischargable"] = false;
                Ticket.hil_JobClosuredon = (DateTime)DateTime.Now.AddMinutes(330);
                //Ticket.msdyn_TimeClosed = (DateTime)DateTime.Now.AddMinutes(330);
                service.Update(Ticket);
            }
        }
        public static Money CalculateChargesInWarranty(IOrganizationService service, Guid TemplateId, Guid IncidentId)
        {
            Money Total = new Money();
            Money PartTotal = CalculatePartChargesInWarranty(service, TemplateId, IncidentId);
            Money LaborTotal = CalculateLaborChargesInWarranty(service, TemplateId, IncidentId);
            decimal Tot = PartTotal.Value + LaborTotal.Value;
            Total = new Money(Tot);
            return Total;
        }
        public static Money CalculateActualChargesOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            Money Total = new Money();
            Money ServicesTotal = CalculateServicesTotalOutWarranty(service, IncidentId, TicketId);
            Money PartsTotal = CalculatePartsTotalOutWarranty(service, IncidentId, TicketId);
            decimal Tot = ServicesTotal.Value + PartsTotal.Value;
            Total = new Money(Tot);
            return Total;
        }
        public static Money CalculateServicesTotalOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            Money Total = new Money();
            decimal Mid = 0;
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet("msdyn_service", "hil_warrantystatus");
            Qry.AddAttributeValue("msdyn_workorderincident", IncidentId);
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            Qry.AddAttributeValue("msdyn_linestatus", 690970001);
            Qry.AddAttributeValue("hil_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderservice Srvc in Found.Entities)
                {
                    if (Srvc.msdyn_Service != null)
                    {
                        if (Srvc.hil_WarrantyStatus.Value != 1)
                        {
                            Product ThisProduct = (Product)service.Retrieve(Product.EntityLogicalName, Srvc.msdyn_Service.Id, new ColumnSet("hil_amount"));
                            if (ThisProduct.hil_Amount != null)
                            {
                                Money Charge = ThisProduct.hil_Amount;
                                Mid = Mid + Charge.Value;
                            }
                        }
                    }
                }
            }
            Total = new Money(Mid);
            return Total;
        }
        public static Money CalculatePartsTotalOutWarranty(IOrganizationService service, Guid IncidentId, Guid TicketId)
        {
            decimal Total = 0;
            Money TotalAmount = new Money();
            QueryByAttribute Qry = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet("hil_replacedpart", "msdyn_product", "hil_warrantystatus", "msdyn_quantity");
            Qry.AddAttributeValue("msdyn_workorderincident", IncidentId);
            Qry.AddAttributeValue("msdyn_workorder", TicketId);
            Qry.AddAttributeValue("hil_linestatus", 910590000);
            Qry.AddAttributeValue("hil_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct WoPdt in Found.Entities)
                {
                    if (WoPdt.msdyn_Product != null)
                    {
                        if (WoPdt.hil_WarrantyStatus.Value != 1)
                        {
                            Product ThisProduct = (Product)service.Retrieve(Product.EntityLogicalName, WoPdt.hil_replacedpart.Id, new ColumnSet("hil_amount"));
                            if (ThisProduct.hil_Amount != null && WoPdt.hil_replacedpart != null)
                            {
                                decimal CalculatedCharge = Convert.ToDecimal(ThisProduct.hil_Amount.Value * Convert.ToDecimal(WoPdt.msdyn_Quantity));
                                Money Charge = new Money(CalculatedCharge);
                                Total = Total + Charge.Value;
                            }
                        }
                    }
                }
            }
            TotalAmount = new Money(Total);
            return TotalAmount;
        }
        public static Money CalculatePartChargesInWarranty(IOrganizationService service, Guid TemplateId, Guid IncidentId)
        {
            decimal Total = 0;
            Money total = new Money();
            QueryByAttribute Query = new QueryByAttribute(hil_part.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_warrantytemplateid", TemplateId);
            Query.AddAttributeValue("hil_includedinwarranty", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (hil_part Part in Found.Entities)
                {
                    if (Part.hil_PartCode != null)
                    {
                        bool CheckIfEstimated = CheckIfPartUsed(service, Part.hil_PartCode.Id, IncidentId);
                        if (CheckIfEstimated == true)
                        {
                            decimal GetQuantity = GetQuantityFromJobProduct(service, Part.hil_PartCode.Id, IncidentId);
                            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Part.hil_PartCode.Id, new ColumnSet("hil_amount"));
                            Total = Total + (Pdt.hil_Amount.Value * GetQuantity);
                        }
                    }
                }
            }
            total = new Money(Total);
            return total;
        }
        public static decimal GetQuantityFromJobProduct(IOrganizationService service, Guid Part, Guid Incident)
        {
            decimal Quantity = 1;
            QueryByAttribute query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            query.ColumnSet = new ColumnSet("msdyn_quantity");
            query.AddAttributeValue("msdyn_workorderincident", Incident);
            query.AddAttributeValue("msdyn_product", Part);
            query.AddAttributeValue("msdyn_linestatus", 690970001);
            EntityCollection Found = service.RetrieveMultiple(query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct WoPdt in Found.Entities)
                {
                    Quantity = Convert.ToDecimal(WoPdt.msdyn_Quantity);
                }
            }
            return Quantity;
        }
        public static Money CalculateLaborChargesInWarranty(IOrganizationService service, Guid TemplateId, Guid IncidentId)
        {
            decimal Total = 0;
            Money total = new Money();
            QueryByAttribute Query = new QueryByAttribute(hil_labor.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_warrantytemplateid", TemplateId);
            Query.AddAttributeValue("hil_includedinwarranty", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (hil_labor Labor in Found.Entities)
                {
                    if (Labor.hil_Labor != null)
                    {

                        bool CheckIfEstimated = CheckIfLaborUsed(service, Labor.hil_Labor.Id, IncidentId);
                        if (CheckIfEstimated == true)
                        {
                            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Labor.hil_Labor.Id, new ColumnSet("hil_amount"));
                            Total = Total + Pdt.hil_Amount.Value;
                        }
                    }
                }
            }
            total = new Money(Total);
            return total;
        }
        public static bool CheckIfPartUsed(IOrganizationService service, Guid Part, Guid Incident)
        {
            bool MarkIfUsed = false;
            QueryByAttribute query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            query.ColumnSet = new ColumnSet(false);
            query.AddAttributeValue("msdyn_workorderincident", Incident);
            query.AddAttributeValue("msdyn_product", Part);
            query.AddAttributeValue("msdyn_linestatus", 690970001);
            EntityCollection Found = service.RetrieveMultiple(query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct Pdt in Found.Entities)
                {
                    if (Pdt.hil_replacedpart != null)
                    {
                        MarkIfUsed = true;
                    }
                }
            }
            return MarkIfUsed;
        }
        public static bool CheckIfLaborUsed(IOrganizationService service, Guid Labor, Guid Incident)
        {
            bool MarkIfUsed = false;
            msdyn_workorderservice Serv = new msdyn_workorderservice();
            QueryByAttribute query = new QueryByAttribute(msdyn_workorderservice.EntityLogicalName);
            query.ColumnSet = new ColumnSet(false);
            query.AddAttributeValue("msdyn_workorderincident", Incident);
            query.AddAttributeValue("msdyn_service", Labor);
            query.AddAttributeValue("msdyn_linestatus", 690970001);
            EntityCollection Found = service.RetrieveMultiple(query);
            if (Found.Entities.Count >= 1)
            {
                MarkIfUsed = true;
            }
            return MarkIfUsed;
        }
        #endregion
    }
}
