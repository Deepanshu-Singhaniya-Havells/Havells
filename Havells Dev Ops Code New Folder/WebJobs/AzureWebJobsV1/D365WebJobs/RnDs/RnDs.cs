using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using RestSharp;

namespace D365WebJobs
{
    public class RnDs
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
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {

                //AssignAMCInvoiceTBSH.AssignRecord(_service); 

                WarrantyEngine.ExecuteWarrantyEngine(_service, new Guid("C36D3B15-BD34-ED11-9DB2-6045BDAC5348"));
                Console.WriteLine("Done");
                //SendEmailOnActivityComplete(new Entity("hil_grievancehandlingactivity", new Guid("1ba0f440-fd10-ef11-9f89-6045bde83873")), _service);

                //EntityReference caseEntityReference = new EntityReference("incident", new Guid("0e44e3ea-ca00-4e62-a035-8202ba675731"));
                //// First close the Incident

                //// Create resolution for the closing incident
                //IncidentResolution incidentRresolution = new IncidentResolution
                //{
                //    Subject = "Resolution Test"
                //};

                //incidentRresolution.IncidentId = caseEntityReference;

                //// Create the request to close the incident, and set its resolution to the resolution created above
                //CloseIncidentRequest closeIncidentRequest = new CloseIncidentRequest();
                //closeIncidentRequest.IncidentResolution = incidentRresolution;

                //// Set the requested new status for the closed Incident
                //closeIncidentRequest.Status = new OptionSetValue(5 /*ProblemSolved*/);

                //// Execute the close request
                //CloseIncidentResponse closeIncidentResponse = (CloseIncidentResponse)_service.Execute(closeIncidentRequest);
                //Console.WriteLine("Resolved.");

                //HavellsOneWebsiteAMCPlanData(_service);
                //Entity _entRMCostSheetLine = _service.Retrieve("hil_rmcostsheetline", new Guid("EB058EDF-32FE-EE11-9F89-7C1E5205F266"), new ColumnSet("hil_rmcostsheet"));

                //CalculateRMCost(_service, _entRMCostSheetLine.GetAttributeValue<EntityReference>("hil_rmcostsheet").Id);

                //string  _api = "https://japi.instaalerts.zone/failsafe/HttpLink?aid=640990&pin=w~7Xg)9V&mnumber=8285906486&signature=HAVELL&message=Welcome Kuldeep, You are now our registered Consumer! Install Havells Sync App for product experience and service at fingertips http://onelink.to/3sdqa7 - Havells&dlt_entity_id=110100001483&dlt_template_id=1107167455882315674&cust_ref=Kuldeep";
                //QueryExpression query = new QueryExpression("hil_smsconfiguration");
                //query.ColumnSet = new ColumnSet("hil_message", "createdon", "regardingobjectid", "hil_mobilenumber", "hil_smstemplate");
                //query.Criteria.AddCondition("hil_smstemplate", ConditionOperator.Equal, technicianTemplate);
                //query.Criteria.AddCondition("createdon", ConditionOperator.Today);

                //Entity entity1 = _service.Retrieve("",new Guid(),new ColumnSet(true));


                //Entity entity = _service.Retrieve("hil_smsconfiguration", new Guid("2b4f7d14-98f5-ee11-a1fe-6045bde78af4"), new ColumnSet(true));
                //hil_smsconfiguration SMS = entity.ToEntity<hil_smsconfiguration>();
                //Guid _regardingObjectId = Guid.Empty;

                //if (entity.Contains("hil_smstemplate"))
                //{
                //    if (entity.Attributes.Contains("regardingobjectid"))
                //    {
                //        _regardingObjectId = entity.GetAttributeValue<EntityReference>("regardingobjectid").Id;
                //    }

                //    EntityReference _entSMSTemplate = entity.GetAttributeValue<EntityReference>("hil_smstemplate");
                //    Entity _smsTemplate = _service.Retrieve("hil_smstemplates", _entSMSTemplate.Id, new ColumnSet("hil_templateid", "hil_encryptsms"));
                //    string _smsTemplateId = _smsTemplate.GetAttributeValue<string>("hil_templateid");
                //    string message = SMS.hil_Message;
                //    message = message.Replace("#", "%23");
                //    message = message.Replace("&", "%26");
                //    message = message.Replace("+", "%2B");

                //    SMS.hil_MobileNumber = "8810570149";

                //    Entity entTemp = new Entity(entity.LogicalName, entity.Id);
                //    if (_smsTemplateId != "1107161191448698079" && _smsTemplateId != "1107161191438154934")
                //    {
                //        SMS = HelperShootSMS.OnDemandSMSShootFunctionDLT(_service, message, SMS.hil_MobileNumber, SMS, _smsTemplateId, SMS.ActivityAdditionalParams, _regardingObjectId);

                //        entTemp["hil_responsefromserver"] = SMS["hil_responsefromserver"];
                //        entTemp["hil_message"] = SMS["hil_message"];
                //        //entTemp["statuscode"] = new OptionSetValue(2); // Sent|Pending
                //        //entTemp["statecode"] = new OptionSetValue(1); // Mark as completed
                //    }
                //    if (_smsTemplate.GetAttributeValue<bool>("hil_encryptsms"))
                //    {
                //        entTemp["hil_message"] = HelperShootSMS.EncryptString(HelperShootSMS.EncriptionKey, (string)entity["hil_message"]);
                //        entTemp["hil_encrypted"] = true;
                //    }
                //    _service.Update(entTemp);

                //    //SetStateRequest req = new SetStateRequest();
                //    //req.State = new OptionSetValue(1);
                //    //req.Status = new OptionSetValue(2);
                //    //req.EntityMoniker = SMS.ToEntityReference();
                //    //SetStateResponse res = (SetStateResponse)service.Execute(req);
                //}

                Console.WriteLine("Hello");
                
                //WebRequest request = WebRequest.Create(_api);
                //request.Method = "POST";
                //WebResponse response = null;
                //string IfOkay = string.Empty;
                //string responseFromServer = string.Empty;
                //try
                //{
                //    response = request.GetResponse();
                //    Stream dataStream = Stream.Null;
                //    IfOkay = ((HttpWebResponse)response).StatusDescription;
                //    dataStream = response.GetResponseStream();
                //    StreamReader reader = new StreamReader(dataStream);
                //    responseFromServer = reader.ReadToEnd();
                //}
                //catch (WebException ex)
                //{
                //    Console.WriteLine(ex.Message);
                //}

                //Guid _ownerId = Guid.Empty;
                //DateTime _resolveBy = new DateTime(2024, 3, 6, 14, 35, 0);
                //TimeSpan _diffMin = DateTime.Now - _resolveBy;
                //string  remarks;
                //if (_diffMin.TotalMinutes <= 60)
                //{
                //    remarks = Convert.ToInt32(_diffMin.TotalMinutes) + " Minutes";
                //}
                //else if (_diffMin.TotalMinutes < 1440)
                //{
                //    double _hrs = _diffMin.TotalMinutes / 60.0;
                //    if (Convert.ToInt32(_diffMin.TotalMinutes) % 60 != 0)
                //    {
                //        remarks = Convert.ToInt32(_diffMin.TotalMinutes)/60 + 1 + " hrs";
                //    }
                //    else
                //        remarks = Math.Round(_hrs, 0) + " hrs";
                //}
                //else
                //{
                //    int _hr = Convert.ToInt32(_diffMin.TotalMinutes) / 1440;
                //    remarks = _hr + " days";
                //}
                //Console.WriteLine(remarks);
                Console.ReadLine();
                //CDR_Request1 _request = new CDR_Request1()
                //{
                //    Caller_Id = "8971930432",
                //    Caller_Name = "Deepanshu Singhaniya",
                //    Caller_Number = "8285906486",
                //    Caller_Status = "Disconnected",
                //    Call_Type = "OUTBOUND",
                //    Conversation_Duration = "00:00:18",
                //    Correlation_ID = "Xchangeec558ce8-dd3c-448c-b115-4b34171e2ebb",
                //    Date = "21/09/2023",
                //    Destination_Name = "Deepanshu Singhaniya",
                //    Destination_Number = "8810570149",
                //    Destination_Status = "Disconnected",
                //    Overall_Call_Duration = "00:00:27",
                //    Overall_Call_Status = "Answered",
                //    Recording = "https://openapi.airtel.in/gateway/airtel-xchange/ironman-data-transfer/download/recordingFile?token=gLTIM5uvaO6TqDKclWycfdBfmt7769wWxG64HnwA035a16Wyuh8aQpEu8xTGLfNNjxu3Oj2CaJBtvaPniEymdo7Hmls+8Y2hWF6DFD0VZMc=",
                //    Time = "09:30:55"
                //};
                //GrivanceModule _CDR = new GrivanceModule();
                //_CDR.CreateCallCDRREcord(_service, _request);
                //Console.ReadLine();
                //GrivanceModule.GrivanceModule_PostCreate(_service);
                //string _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //    <entity name='hil_caseassignmentmatrix'>
                //    <attribute name='hil_caseorigin' />
                //    <attribute name='hil_department' />
                //    <attribute name='hil_spocsla' />
                //    <attribute name='hil_firstresponsesla' />
                //    <attribute name='hil_caseassignmentmatrixid' />
                //    <attribute name='hil_name' />
                //    <order attribute='hil_caseorigin' descending='false' />
                //    <filter type='and'>
                //        <condition attribute='statecode' operator='eq' value='0' />
                //        <condition attribute='hil_department' operator='eq' uiname='Service' uitype='hil_casedepartment' value='{AB3DBC3D-4E6E-EE11-8179-6045BDAC526A}' />
                //    </filter>
                //    </entity>
                //    </fetch>";

                //EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchxml));
                //foreach (Entity ent in entCol.Entities) {
                //    Console.WriteLine(ent.GetAttributeValue<String>("hil_name") + " :FirstResSLA " + ent.GetAttributeValue<Int32>("hil_firstresponsesla") + " :SPOC SLA " + ent.GetAttributeValue<Int32>("hil_spocsla"));
                //}
                //Console.ReadLine();
                //Entity entity = _service.Retrieve(entityPhoneCall.LogicalName, entityPhoneCall.Id, new ColumnSet("ownerid", "regardingobjectid", "hil_contactpreference", "hil_callingnumber"));

                //if (entity.Contains("regardingobjectid"))
                //{
                //    JobID = entity.GetAttributeValue<EntityReference>("regardingobjectid").Name;
                //}
                //if (entity.Contains("ownerid"))
                //{
                //    _ownerId = entity.GetAttributeValue<EntityReference>("ownerid").Id;
                //}
                //if (entity.Contains("hil_contactpreference"))
                //{
                //    preference = entity.GetAttributeValue<OptionSetValue>("hil_contactpreference").Value;
                //    isPreference = true;
                //}
                //if (entity.Contains("hil_callingnumber"))
                //{
                //    callingNumber = entity.GetAttributeValue<string>("hil_callingnumber");
                //}

                //if (!string.IsNullOrEmpty(JobID) && entity.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName == "msdyn_workorder" && isPreference)
                {
                    //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    //      <entity name='systemuser'>
                    //        <attribute name='systemuserid' />
                    //        <filter type='and'>
                    //          <condition attribute='systemuserid' operator='eq' value='{_ownerId}' />
                    //        </filter>
                    //        <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                    //          <link-entity name='role' from='roleid' to='roleid' alias='ae'>
                    //            <filter type='and'>
                    //              <condition attribute='name' operator='eq' value='{HelperClass._callMaskingRoleName}' />
                    //            </filter>
                    //          </link-entity>
                    //        </link-entity>
                    //      </entity>
                    //    </fetch>";
                    //EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    //if (_entCol.Entities.Count > 0)
                    //{
                    //    //string result = C2CAPI(service, entity);
                    //}
                }

                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //      <entity name='contact'>
                //        <attribute name='createdon' />
                //        <attribute name='contactid' />
                //        <attribute name='hil_consumersource' />
                //        <attribute name='mobilephone' />
                //        <attribute name='middlename' />
                //        <attribute name='firstname' />
                //        <order attribute='createdon' descending='true' />
                //        <filter type='and'>
                //          <filter type='or'>
                //            <condition attribute='hil_consumersource' operator='null' />
                //            <condition attribute='fullname' operator='null' />
                //          </filter>
                //          <condition attribute='mobilephone' operator='not-null' />
                //        </filter>
                //      </entity>
                //    </fetch>";
                //EntityCollection entCol = null;
                //int i = 1,rowCnt = 0;
                //while (true)
                //{
                //    entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //    rowCnt += entCol.Entities.Count;

                //    if (entCol.Entities.Count == 0) { break; }
                //    foreach (Entity ent in entCol.Entities)
                //    {
                //        DateTime _createdOn = ent.GetAttributeValue<DateTime>("createdon");
                //        Entity entUpdate = new Entity(ent.LogicalName, ent.Id);
                //        if (!ent.Contains("firstname"))
                //            entUpdate["firstname"] = ent.GetAttributeValue<string>("mobilephone");
                //        if (!ent.Contains("hil_consumersource"))
                //            entUpdate["hil_consumersource"] = new OptionSetValue(1);
                //        _service.Update(entUpdate);
                //        Console.WriteLine("Processing... " + ent.GetAttributeValue<string>("mobilephone") + " /"+ _createdOn.ToString() + " :: "+ i.ToString() + "/" + rowCnt.ToString());
                //        i++;
                //    }
                //}

                //Money _receiptAmount = new Money(new decimal(4500.00));
                //Console.WriteLine(_receiptAmount.Value.ToString("0.00"));
                //Console.ReadLine();
                //Entity _ent = _service.Retrieve("msdyn_customerasset", new Guid("83794d2f-daa0-eb11-b1ac-6045bd72c485"), new ColumnSet("hil_invoicedate"));

                //DateTime _dop = _ent.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330);
                //int age = Convert.ToInt32(((DateTime.Now- _dop).TotalDays / 365));
                //string AssetAge = age < 1 ? "<1 yr" : age + " yr";
                //string DOP = _ent.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToString("yyyy-MMMM-dd");

                //PostCreateAsync_PopulateSpares _restrictSparePart = new PostCreateAsync_PopulateSpares();
                //_restrictSparePart.Execute(_service);
                //Console.WriteLine("DONE");
                //PublishProductHierarchyRequest publishReq = new PublishProductHierarchyRequest
                //{
                //    Target = new EntityReference(Product.EntityLogicalName, new Guid("5b09ff0a-48a6-ea11-a812-000d3af0563c"))
                //};
                //PublishProductHierarchyResponse published = (PublishProductHierarchyResponse)_service.Execute(publishReq);
                //if (published.Results != null)
                //{
                //    Console.WriteLine("Published the product rows");
                //}

                //Entity caseEntity = _service.Retrieve("incident", new Guid("30dd25a6-c977-4693-8f3d-ccff00150baa"), new ColumnSet(true));
                //int caseOrigin = 0;
                //EntityReference entCaseCategory = null;
                //EntityReference entCaseProduct = null;
                //int caseType = caseEntity.GetAttributeValue<OptionSetValue>("casetypecode").Value;
                //EntityReference entDepartment = caseEntity.GetAttributeValue<EntityReference>("hil_casedepartment");
                //EntityReference entBranch = caseEntity.GetAttributeValue<EntityReference>("hil_branch");
                //string _fetchXML = string.Empty;
                //EntityCollection _entColl = null;
                //if (entDepartment.Id == new Guid("7bf1705a-3764-ee11-8df0-6045bdaa91c3"))//Sampark Call Center and Case Type !=Enquiry. For Ennquiry we are resolving Case On Call.
                //{
                //    if (caseType != 1)
                //    {
                //        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //              <entity name='hil_caseassignmentmatrixline'>
                //                <attribute name='hil_assigneeuser' />
                //                <attribute name='hil_assigneeteam' />
                //                <attribute name='hil_caseassignmentmatrixid' />
                //                <filter type='and'>
                //                    <condition attribute='statecode' operator='eq' value='0' />
                //                    <condition attribute='hil_level' operator='eq' value='1' />
                //                    <filter type='or'>
                //                        <condition attribute='hil_assigneeuser' operator='not-null' />
                //                        <condition attribute='hil_assigneeteam' operator='not-null' />
                //                    </filter>
                //                </filter>
                //                <link-entity name='hil_caseassignmentmatrix' from='hil_caseassignmentmatrixid' to='hil_caseassignmentmatrixid' link-type='inner' alias='ac'>
                //                  <filter type='and'>
                //                    <condition attribute='hil_department' operator='eq' value='{entDepartment.Id}' />
                //                    <condition attribute='hil_branch' operator='eq' value='{entBranch.Id}' />
                //                    <condition attribute='statecode' operator='eq' value='0' />
                //                  </filter>
                //                </link-entity>
                //              </entity>
                //            </fetch>";
                //        _entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //        if (_entColl.Entities.Count > 0)
                //        {
                //            EntityReference _assignmentMatrix = _entColl[0].GetAttributeValue<EntityReference>("hil_caseassignmentmatrixid");
                //            EntityReference _assignee = _entColl[0].Contains("hil_assigneeuser") ? _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeuser") : _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                //            AssignCase(_service, caseEntity, _assignee, _assignmentMatrix);
                //        }
                //    }
                //}
                //else
                //{
                //    caseOrigin = caseEntity.GetAttributeValue<OptionSetValue>("caseorigincode").Value;
                //    entCaseCategory = caseEntity.GetAttributeValue<EntityReference>("hil_casecategory");
                //    entCaseProduct = caseEntity.GetAttributeValue<EntityReference>("productid");

                //    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //        <entity name='hil_caseassignmentmatrixline'>
                //        <attribute name='hil_caseassignmentmatrixlineid' />
                //        <attribute name='hil_assigneeuser' />
                //        <attribute name='hil_assigneeteam' />
                //        <filter type='and'>
                //                <condition attribute='statecode' operator='eq' value='0' />
                //                <condition attribute='hil_level' operator='eq' value='1' />
                //                <filter type='or'>
                //                    <condition attribute='hil_assigneeuser' operator='not-null' />
                //                    <condition attribute='hil_assigneeteam' operator='not-null' />
                //                </filter>
                //            </filter>
                //        <link-entity name='hil_caseassignmentmatrix' from='hil_caseassignmentmatrixid' to='hil_caseassignmentmatrixid' link-type='inner' alias='af'>
                //        <filter type='and'>
                //        <condition attribute='hil_department' operator='eq' value='{entDepartment.Id}' />
                //        <condition attribute='hil_branch' operator='eq' value='{entBranch.Id}' />
                //        <condition attribute='statecode' operator='eq' value='0' />
                //        <condition attribute='hil_caseorigin' operator='eq' value='{caseOrigin}' />
                //        <condition attribute='hil_casetype' operator='eq' value='{caseType}' />
                //        <condition attribute='hil_casecategory' operator='eq' value='{entCaseCategory.Id}' />";
                //    if (entCaseProduct != null)
                //        _fetchXML = _fetchXML + $@"<condition attribute='hil_productdivision' operator='eq' value='{entCaseProduct.Id}' />";
                //    _fetchXML = _fetchXML + $@"</filter>
                //        </link-entity>
                //        </entity>
                //        </fetch>";
                //    _entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //    if (_entColl.Entities.Count > 0)
                //    {
                //        EntityReference _assignmentMatrix = _entColl[0].ToEntityReference();
                //        EntityReference _assignee = _entColl[0].Contains("hil_assigneeuser") ? _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeuser") : _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                //        AssignCase(_service, caseEntity, _assignee, _assignmentMatrix);
                //    }
                //}

                //if (caseType != 1 && entDepartment.Id == new Guid("7bf1705a-3764-ee11-8df0-6045bdaa91c3"))//Sampark Call Center and Case Type !=Enquiry. For Ennquiry we are
                //{
                //    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //          <entity name='hil_caseassignmentmatrixline'>
                //            <attribute name='hil_assigneeuser' />
                //            <attribute name='hil_assigneeteam' />
                //            <filter type='and'>
                //              <condition attribute='hil_level' operator='eq' value='1' />
                //              <condition attribute='statecode' operator='eq' value='0' />
                //            </filter>
                //            <link-entity name='hil_caseassignmentmatrix' from='hil_caseassignmentmatrixid' to='hil_caseassignmentmatrixid' link-type='inner' alias='ac'>
                //              <filter type='and'>
                //                <condition attribute='hil_department' operator='eq' value='{entDepartment.Id}' />
                //                <condition attribute='hil_branch' operator='eq' value='{entBranch.Id}' />
                //                <condition attribute='statecode' operator='eq' value='0' />
                //              </filter>
                //            </link-entity>
                //          </entity>
                //        </fetch>";
                //    EntityCollection _entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //    if (_entColl.Entities.Count > 0)
                //    {
                //        if (_entColl[0].Contains("hil_assigneeteam"))
                //        {
                //            _entCase["ownerid"] = _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                //            _service.Update(_entCase);
                //        }
                //        else if (_entColl[0].Contains("hil_assigneeuser"))
                //        {
                //            _entCase["ownerid"] = _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeuser");
                //            _service.Update(_entCase);
                //        }
                //    }
                //}

                //int caseOrigin = _entCase.GetAttributeValue<OptionSetValue>("caseorigincode").Value;
                //EntityReference entCaseCategory = _entCase.GetAttributeValue<EntityReference>("hil_casecategory");
                //EntityReference entCaseProduct = _entCase.GetAttributeValue<EntityReference>("productid");

                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //        <entity name='hil_caseassignmentmatrixline'>
                //        <attribute name='hil_caseassignmentmatrixlineid' />
                //        <attribute name='hil_assigneeuser' />
                //        <attribute name='hil_assigneeteam' />
                //        <filter type='and'>
                //        <condition attribute='hil_level' operator='eq' value='1' />
                //        <condition attribute='statecode' operator='eq' value='0' />
                //        </filter>
                //        <link-entity name='hil_caseassignmentmatrix' from='hil_caseassignmentmatrixid' to='hil_caseassignmentmatrixid' link-type='inner' alias='af'>
                //        <filter type='and'>
                //        <condition attribute='hil_department' operator='eq' value='{entDepartment.Id}' />
                //        <condition attribute='hil_branch' operator='eq' value='{entBranch.Id}' />
                //        <condition attribute='statecode' operator='eq' value='0' />
                //        <condition attribute='hil_caseorigin' operator='eq' value='{caseOrigin}' />
                //        <condition attribute='hil_casetype' operator='eq' value='{caseType}' />
                //        <condition attribute='hil_casecategory' operator='eq' value='{entCaseCategory.Id}' />";
                //if (entCaseProduct != null)
                //    _fetchXML = _fetchXML + $@"<condition attribute='hil_productdivision' operator='eq' value='{entCaseProduct.Id}' />";
                //_fetchXML = _fetchXML + $@"</filter>
                //        </link-entity>
                //        </entity>
                //        </fetch>";
                //EntityCollection _entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //if (_entColl.Entities.Count > 0)
                //{
                //    if (_entColl[0].Contains("hil_assigneeteam"))
                //    {
                //        _entCase["ownerid"] = _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                //        _service.Update(_entCase);
                //    }
                //    else if (_entColl[0].Contains("hil_assigneeuser"))
                //    {
                //        _entCase["ownerid"] = _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeuser");
                //        _service.Update(_entCase);
                //    }
                //}

                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //      <entity name='hil_warrantytemplate'>
                //        <attribute name='hil_name' />
                //        <attribute name='createdon' />
                //        <attribute name='hil_warrantyperiod' />
                //        <attribute name='hil_type' />
                //        <attribute name='hil_product' />
                //        <attribute name='hil_validto' />
                //        <attribute name='hil_validfrom' />
                //        <attribute name='hil_templatestatus' />
                //        <attribute name='hil_isapproved' />
                //        <attribute name='hil_amcplan' />
                //        <attribute name='hil_warrantyapplicationon' />
                //        <attribute name='hil_warrantytemplateid' />
                //        <attribute name='hil_applicableon' />
                //        <order attribute='createdon' descending='true' />
                //        <filter type='and'>
                //          <condition attribute='statecode' operator='eq' value='0' />
                //          <condition attribute='hil_applicableon' operator='not-null' />
                //          <condition attribute='hil_warrantyapplicationon' operator='null' />
                //        </filter>
                //      </entity>
                //    </fetch>";
                //EntityCollection entColOrderBy = null;
                //int i = 1, totalRecords = 0;
                //while (true)
                //{
                //    entColOrderBy = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //    totalRecords += entColOrderBy.Entities.Count;
                //    foreach (Entity ent in entColOrderBy.Entities)
                //    {
                //        Entity _entUpdate = new Entity(ent.LogicalName,ent.Id);
                //        _entUpdate["hil_warrantyapplicationon"] = new EntityReference("hil_warrantytemplateexecutionindex", new Guid("80ECAC41-FFA2-EE11-A569-6045BDAC5778"));
                //        _service.Update(_entUpdate);
                //        Console.WriteLine("Updating ... " + i++ + "/" + totalRecords);
                //    }
                //}
                //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //      <entity name='hil_warrantytemplate'>
                //        <attribute name='hil_name' />
                //        <attribute name='createdon' />
                //        <attribute name='hil_warrantyperiod' />
                //        <attribute name='hil_type' />
                //        <attribute name='hil_product' />
                //        <attribute name='hil_category' />
                //        <attribute name='hil_warrantytemplateid' />
                //        <order attribute='createdon' descending='false' />
                //        <filter type='and'>
                //          <condition attribute='hil_templatestatus' operator='eq' value='2' />
                //          <condition attribute='hil_applicableon' operator='eq' value='910590000' />
                //          <condition attribute='hil_type' operator='eq' value='2' />
                //        </filter>
                //        <link-entity name='product' from='productid' to='hil_product' visible='false' link-type='outer' alias='a_ad77818aea04e911a94d000d3af06c56'>
                //          <attribute name='createdon' />
                //        </link-entity>
                //      </entity>
                //    </fetch>";
                //EntityCollection entColOrderBy = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //if (entColOrderBy.Entities.Count > 0)
                //{
                //    Console.WriteLine(entColOrderBy.Entities[0].GetAttributeValue<string>("hil_name"));
                //}

                //UpdateJobWarrantyDetails(new Guid ("e440b861-bc56-ee11-be6f-6045bdac526a"), _service);
                //HelperWarrantyModule.Init_Warranty(new Guid("cb09b6ec-5a9a-ee11-be37-000d3a3e4841"), _service);

                //string _fetchXML = @" 
                //<fetch distinct='false' mapping='logical' aggregate='true'> 
                //   <entity name='opportunity'> 
                //      <attribute name='name' alias='opportunity_count' aggregate='countcolumn' /> 
                //      <attribute name='ownerid' alias='ownerid' groupby='true' /> 
                //   </entity> 
                //</fetch>";

                //CreateEnquiry _object = new CreateEnquiry();
                //ReturnInfoVar _retObj = _object.CreateEnquiryEntry(new CreateEnquiry() { 
                //CampaignCode= "CAMP2000006",
                //Email="Azad.kumar@havells.com",
                //EnquiryType= "910590001",
                //MobNumber="8287910060",
                //Name = "Azad Kumar",
                //Pincode="201304",
                //productDivisionID="44,45,44,45",
                //Remarks="Test Entry"
                //}, _service);

                //try
                //{
                //    string _jobIncidentId = "54cdd144-985e-ed11-9562-6045bdac5292";
                //    string _serviceBOMId = "11f2f2af-c3d9-ed11-a7c7-6045bdac5098";

                //    if (string.IsNullOrEmpty(_jobIncidentId) || string.IsNullOrEmpty(_serviceBOMId))
                //    {
                //        Console.WriteLine("204|Job Incident Id and Service BOM Id is required.");
                //    }
                //    else
                //    {
                //        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //            <entity name='msdyn_workorderincident'>
                //            <attribute name='msdyn_workorderincidentid' />
                //            <attribute name='msdyn_workorder' />
                //            <attribute name='msdyn_customerasset' />
                //            <filter type='and'>
                //            <condition attribute='msdyn_workorderincidentid' operator='eq' value='{_jobIncidentId}' />
                //            <condition attribute='statecode' operator='eq' value='0' />
                //            </filter>
                //            </entity>
                //            </fetch>";

                //        EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //        if (entCol.Entities.Count > 0)
                //        {
                //            try
                //            {
                //                Entity _serviceBOM = _service.Retrieve("hil_servicebom", new Guid(_serviceBOMId), new ColumnSet("hil_isserialized", "hil_quantity", "hil_product", "hil_priority", "hil_chargeableornot"));
                //                if (_serviceBOM != null)
                //                {

                //                    Entity _jobProduct = new Entity("msdyn_workorderproduct");
                //                    _jobProduct["msdyn_customerasset"] = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");

                //                    if (_serviceBOM.Contains("hil_product"))
                //                    {
                //                        EntityReference _sparePart = _serviceBOM.GetAttributeValue<EntityReference>("hil_product");
                //                        _jobProduct["msdyn_product"] = _sparePart;
                //                        Product Pdt1 = (Product)_service.Retrieve(Product.EntityLogicalName, _sparePart.Id, new ColumnSet("name", "description", "hil_amount"));
                //                        string Uq = Pdt1.Description != null ? _sparePart.Name + "-" + Pdt1.Description : _sparePart.Name + "-";
                //                        _jobProduct["hil_part"] = Uq;
                //                        _jobProduct["hil_priority"] = _serviceBOM.Contains("hil_priority") ? (string)_serviceBOM["hil_priority"] : string.Empty;

                //                        _jobProduct["msdyn_workorderincident"] = entCol.Entities[0].ToEntityReference();

                //                        _jobProduct["msdyn_workorder"] = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_workorder");

                //                        _jobProduct["hil_maxquantity"] = Convert.ToDecimal(_serviceBOM.Contains("hil_quantity") ? _serviceBOM.GetAttributeValue<int>("hil_quantity") : 1);
                //                        _jobProduct["msdyn_quantity"] = Convert.ToDouble(1);

                //                        if (_serviceBOM.Contains("hil_chargeableornot"))
                //                        {
                //                            OptionSetValue _chargeableOS = _serviceBOM.GetAttributeValue<OptionSetValue>("hil_chargeableornot");
                //                            _jobProduct["hil_chargeableornot"] = _chargeableOS;
                //                            if (_chargeableOS.Value == 1)
                //                            {
                //                                _jobProduct["hil_warrantystatus"] = new OptionSetValue(2);
                //                            }
                //                        }
                //                        if (Pdt1.hil_Amount != null)
                //                        {
                //                            _jobProduct["msdyn_totalamount"] = Pdt1.hil_Amount;
                //                            _jobProduct["1hil_partamount"] = Pdt1.hil_Amount.Value;
                //                        }
                //                        if (_serviceBOM.Contains("hil_isserialized"))
                //                        {
                //                            _jobProduct["hil_isserialized"] = _serviceBOM.GetAttributeValue<OptionSetValue>("hil_isserialized");
                //                        }
                //                    }
                //                    _service.Create(_jobProduct);
                //                }
                //                Console.WriteLine("200|OK.");
                //            }
                //            catch (Exception ex)
                //            {
                //                Console.WriteLine("204|" + ex.Message);
                //            }
                //        }
                //        else
                //        {
                //            Console.WriteLine("204|Job Incident Id does not exist.");
                //        }
                //    }
                //}
                //catch (Exception ex)
                //{
                //    throw new InvalidPluginExecutionException(ex.Message);
                //}

                //SetStateRequest req = new SetStateRequest();
                //req.State = new OptionSetValue(1);
                //req.Status = new OptionSetValue(2);
                //req.EntityMoniker = new EntityReference("hil_smsconfiguration", new Guid("136abceb-db72-ee11-8179-6045bda59a94"));
                //SetStateResponse res = (SetStateResponse)_service.Execute(req);

                //ServiceBOMResponse obj = JobServiceProductAction.PopulateSpareParts(new Guid("54cdd144-985e-ed11-9562-6045bdac5292"), _service);
                //Console.WriteLine(obj._partList.Count.ToString());

                //Entity _enJob1 = _service.Retrieve("msdyn_workorder", new Guid("88fc4919-1ac7-ed11-b597-6045bdac54b1"), new ColumnSet("createdon", "hil_customerref", "hil_callsubtype", "hil_productsubcategory", "msdyn_customerasset"));
                //DateTime _createdOn = _enJob1.GetAttributeValue<DateTime>("createdon").AddDays(-15);
                //DateTime _ClosedOn = DateTime.Now.AddDays(-15);
                //string _strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
                //string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();
                //EntityCollection entCol;

                ////Callsubtype{8D80346B-3C0B-E911-A94E-000D3AF06CD} Dealer Stock Repair
                ////JobStatus: {1727FA6C-FA0F-E911-A94E-000D3AF060A1}-Closed,2927FA6C-FA0F-E911-A94E-000D3AF060A1-Workdone,7E85074C-9C54-E911-A951-000D3AF0677F-workdone SMS
                //string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //    <entity name='msdyn_workorder'>
                //    <attribute name='msdyn_name' />
                //    <attribute name='createdon' />
                //    <attribute name='hil_productsubcategory' />
                //    <attribute name='hil_customerref' />
                //    <attribute name='msdyn_customerasset' />
                //    <attribute name='hil_callsubtype' />
                //    <attribute name='msdyn_workorderid' />
                //    <attribute name='msdyn_timeclosed' />
                //    <attribute name='msdyn_closedby' />
                //    <order attribute='createdon' descending='true' />
                //    <filter type='and'>
                //        <condition attribute='hil_isocr' operator='ne' value='1' />
                //        <condition attribute='hil_typeofassignee' operator='ne' value='{7D1ECBAB-1208-E911-A94D-000D3AF0694E}' />
                //        <condition attribute='msdyn_workorderid' operator='ne' value='" + _enJob1.Id + @"' />
                //        <condition attribute='hil_customerref' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                //        <condition attribute='hil_callsubtype' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                //        <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                //        <condition attribute='hil_productsubcategory' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                //        <filter type='or'>
                //            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                //            <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                //        </filter>
                //        <condition attribute='msdyn_substatus' operator='in'>
                //        <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                //        <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                //        <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                //        </condition>
                //    </filter>
                //    </entity>
                //    </fetch>";
                //EntityCollection entColRepeatRepair = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //if (entColRepeatRepair.Entities.Count > 0)
                //{
                //    string _remarks = string.Empty;
                //    string _remarks1 = string.Empty;
                //    EntityReference _entref = null;
                //    foreach (Entity ent in entColRepeatRepair.Entities)
                //    {
                //        _entref = ent.ToEntityReference();
                //        if (ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id == _enJob1.GetAttributeValue<EntityReference>("msdyn_customerasset").Id)
                //        {
                //            _remarks += ent.GetAttributeValue<string>("msdyn_name") + ",";
                //        }
                //        else
                //        {
                //            _remarks1 += ent.GetAttributeValue<string>("msdyn_name") + ",";
                //        }
                //    }
                //    _remarks = ((_remarks == string.Empty ? "" : "Repeated Jobs with Same Serial Number: " + _remarks + ":\n") + (_remarks1 == string.Empty ? "" : "Repeated Jobs with Same Product Subcategory: " + _remarks1 + ":")).Replace(",:", "");
                //}
                //CommonLib obj = new CommonLib();
                //CommonLib objReturn = obj.CreateSAWActivity(_enJob1.Id, 0, SAWCategoryConst._RepeatRepair, service, _remarks, _entref);
                //if (objReturn.statusRemarks == "OK")
                //{
                //    _underReview = true;
                //}
                //AMCBilling _obj = ValidateAMCReceiptAmount(new AMCBilling() {
                //    JobId = new Guid("602c2e6c-6c54-ee11-be6f-6045bdac5292"),
                //    ReceiptAmount = 4639,
                //    SourceCode = 13
                //}, _service);

                //IoTValidateSerialNumber obj = new IoTValidateSerialNumber();
                //obj.ValidateAssetSerialNumber(new IoTValidateSerialNumber()
                //{
                //    SerialNumber = "67IEG60U00421"
                //}, _service);
                //ITracingService tracingService = null;
                //Entity contactEntity = _service.Retrieve("hil_address", new Guid("6ecc2077-b553-ee11-be6e-6045bdc62d4b"), new ColumnSet(true));
                //Common.CallAPIInsertUpdateAddress(_service, contactEntity, tracingService);
                //string _fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //      <entity name='asyncoperation'>
                //        <attribute name='asyncoperationid' />
                //        <attribute name='createdon' />
                //        <order attribute='createdon' descending='false' />
                //        <filter type='and'>
                //          <condition attribute='statuscode' operator='eq' value='0' />
                //          <condition attribute='createdon' operator='olderthan-x-hours' value='1' />
                //        </filter>
                //      </entity>
                //    </fetch>";
                //while (true)
                //{
                //    EntityCollection entCollJobProd = _service.RetrieveMultiple(new FetchExpression(_fetch));
                //    if (entCollJobProd.Entities.Count == 0) { break; }
                //    if (entCollJobProd.Entities.Count > 0)
                //    {
                //        int i = 1;
                //        foreach (Entity ent in entCollJobProd.Entities)
                //        {
                //            try
                //            {
                //                _service.Delete("asyncoperation", ent.Id);
                //            }
                //            catch (Exception ex)
                //            {
                //                Console.WriteLine(ex.Message);
                //            }
                //            Console.WriteLine("Processing.. " + i++.ToString());
                //        }
                //    }
                //}
                ////QueryExpression expJobProduct = new QueryExpression("msdyn_workorderproduct");
                ////expJobProduct.ColumnSet = new ColumnSet("msdyn_product");
                ////expJobProduct.Criteria = new FilterExpression(LogicalOperator.And);
                ////expJobProduct.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, new Guid("d92e4a82-ff43-ee11-be6f-6045bdac5348"));
                ////expJobProduct.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                ////expJobProduct.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                ////LinkEntity EntityA = new LinkEntity("msdyn_workorderproduct", "hil_productcatalog", "hil_replacedpart", "hil_productcode", JoinOperator.LeftOuter);
                ////EntityA.Columns = new ColumnSet("hil_amctandc");
                ////EntityA.EntityAlias = "Prod";
                ////EntityA.LinkCriteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active Product Catalog Line
                ////expJobProduct.LinkEntities.Add(EntityA);
                ////LinkEntity EntityB = new LinkEntity("msdyn_workorderproduct", "product", "hil_replacedpart", "productid", JoinOperator.Inner);
                ////EntityB.Columns = new ColumnSet(false);
                ////EntityB.LinkCriteria = new FilterExpression(LogicalOperator.And);
                ////EntityB.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001); //AMC Product
                ////EntityB.LinkCriteria.AddCondition("msdyn_fieldserviceproducttype", ConditionOperator.Equal, 690970001); //Non-Inventory
                ////EntityB.EntityAlias = "ProdAMC";
                ////expJobProduct.LinkEntities.Add(EntityB);
                ////EntityCollection entCollJobProd = _service.RetrieveMultiple(expJobProduct);
                ////if (entCollJobProd.Entities.Count == 0)
                ////{
                ////    throw new InvalidPluginExecutionException("Atlease One AMC Product is Required Before Marking Job Close.");
                ////}
                ////else
                ////{
                ////    if (!entCollJobProd.Entities[0].Contains("Prod.hil_amctandc"))
                ////    {
                ////        throw new InvalidPluginExecutionException("Warranty Description is not defined in selected AMC Product Calatog setup.");
                ////    }
                ////}

                ////Entity _jobEnt = _service.Retrieve("msdyn_workorder", new Guid("2ad94292-e44b-ee11-be6f-6045bdaf13b7"), new ColumnSet("hil_isgascharged", "msdyn_substatus", "hil_callsubtype", "hil_productcategory"));
                ////var gasChargeStatus = _jobEnt.GetAttributeValue<Boolean>("hil_isgascharged");
                ////var subStatus = _jobEnt.GetAttributeValue<EntityReference>("msdyn_substatus");
                ////var callSubType = _jobEnt.GetAttributeValue<EntityReference>("hil_callsubtype");
                ////var prodCategory = _jobEnt.GetAttributeValue<EntityReference>("hil_productcategory");

                ////if ((gasChargeStatus) && subStatus.Name == "Work Done" && (callSubType.Id == new Guid("E2129D79-3C0B-E911-A94E-000D3AF06CD4") ||
                ////    callSubType.Id == new Guid("6560565A-3C0B-E911-A94E-000D3AF06CD4") || callSubType.Id == new Guid("8D80346B-3C0B-E911-A94E-000D3AF06CD4"))
                ////    && (prodCategory.Id == new Guid("D51EDD9D-16FA-E811-A94C-000D3AF0694E") || prodCategory.Id == new Guid("2DD99DA1-16FA-E811-A94C-000D3AF06091")))
                ////{
                ////    // do nothing
                ////}
                ////else
                ////{
                ////    throw new InvalidPluginExecutionException(" ***Gas Charge Not Allowed*** ");
                ////}

                ////TechnicianMobileExt _tech = new TechnicianMobileExt();
                ////ConsumablesData _retVal1 = _tech.RequestData(new Request()
                ////{
                ////    JobID = "19082323958227",
                ////    AssetSubcategoryID = "8d1b7022-410b-e911-a94f-000d3af00f43",
                ////    MobileNumber = "9699047429"
                ////}, _service);
                ////ConsumablesData _retVal2 = _tech.RequestData(new Request()
                ////{
                ////    AssetID = "abcd123d",
                ////    JobID = "16082323882799"
                ////}, _service);
                ////Console.WriteLine(_retVal2.AssetAge);
                ////Guid jobGuid = new Guid("982f7446-4734-ee11-bdf4-6045bdac52d3");
                ////OrganizationRequest req = new OrganizationRequest("hil_NewGetPaymentStatusAction");
                ////req["EntityID"] = jobGuid.ToString().Replace("{", "").Replace("}", "");
                ////req["EntityName"] = "msdyn_workorder";
                ////OrganizationResponse response = _service.Execute(req);
                ////string Status = response.Results["Status"].ToString();
                ////string Message = response.Results["Message"].ToString();


                ////CallMasking _reqObj = new CallMasking();
                ////ResposeDataCallMasking _retVal = _reqObj.GetCustomerOpenJobs(_service, new RequestDataCallMasking()
                ////{
                ////    JobNumber = "6668",
                ////    MobileNumber = "9825135212"
                ////});
                ////Console.WriteLine(_retVal.TechnicianMobileNo);

                ////StringBuilder _fetchXML1 = new StringBuilder();
                ////RequestDataCallMasking1 _requestData = new RequestDataCallMasking1() { JobNumber = "", MobileNumber = "8285906486" };

                ////_fetchXML1.Append(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                ////          <entity name='msdyn_workorder'>
                ////            <attribute name='msdyn_name' />
                ////            <attribute name='createdon' />
                ////            <attribute name='hil_customername' />
                ////            <attribute name='msdyn_workorderid' />
                ////            <order attribute='createdon' descending='true' />
                ////            <filter type='and'>
                ////              <condition attribute='hil_mobilenumber' operator='eq' value='" + _requestData.MobileNumber + @"' />
                ////              <condition attribute='msdyn_substatus' operator='not-in'>
                ////                <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                ////                <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                ////                <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{1527FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                ////              </condition>
                ////              <condition attribute='hil_isocr' operator='ne' value='1' />");
                ////if (!string.IsNullOrEmpty(_requestData.JobNumber))
                ////    _fetchXML1.Append(@"<condition attribute='msdyn_name' operator='like' value='%" + _requestData.JobNumber + @"' />");

                ////_fetchXML1.Append(@"</filter>
                ////        <link-entity name='systemuser' from='systemuserid' to='owninguser' visible='false' link-type='outer' alias='user'>
                ////          <attribute name='mobilephone' />
                ////          <attribute name='address1_telephone1' />
                ////        </link-entity>
                ////      </entity>
                ////    </fetch>");

                ////EntityCollection _entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML1.ToString()));
                ////Console.WriteLine(_entColl.Entities.Count.ToString());

                ////TechnicianMobileExt retObj = new TechnicianMobileExt();
                ////retObj.RequestData(_service, new Request() { AssetID = "01HAZGL01629" });
                ////#region Delete Duplicate AMC SAP Invoice Details
                //////string fetchDistinctAsset = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //////    <entity name='hil_refreshjobs'>
                //////    <attribute name='hil_name' />
                //////    <filter type='and'>
                //////        <condition attribute='statecode' operator='eq' value='0' />
                //////    </filter>
                //////    </entity>
                //////    </fetch>";
                //////while (true)
                //////{
                //////    try
                //////    {
                //////        EntityCollection custAssetColl = _service.RetrieveMultiple(new FetchExpression(fetchDistinctAsset));
                //////        if (custAssetColl.Entities.Count == 0) { break; }
                //////        foreach (Entity entAsset in custAssetColl.Entities)
                //////        {
                //////            string _invNo = entAsset.GetAttributeValue<string>("hil_name");
                //////            Console.WriteLine("Processing... " + entAsset.GetAttributeValue<string>("hil_name"));
                //////            string _fetchXML = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                //////                <entity name='hil_amcstaging'>
                //////                <attribute name='hil_serailnumber' alias='srno' groupby='true' /> 
                //////                <attribute name='hil_amcstagingid' alias='cnt' aggregate='count' />
                //////                <filter type='and'>
                //////                    <condition attribute='statecode' operator='eq' value='0' />
                //////                    <condition attribute='hil_name' operator='eq' value='{_invNo}' />
                //////                </filter>
                //////                </entity>
                //////                </fetch>";

                //////            EntityCollection invColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //////            EntityCollection invcntColl = null;
                //////            foreach (Entity ent in invColl.Entities)
                //////            {
                //////                int cnt = Convert.ToInt32(ent.GetAttributeValue<AliasedValue>("cnt").Value);
                //////                string srno = ent.GetAttributeValue<AliasedValue>("srno").Value.ToString();
                //////                if (cnt > 1)
                //////                {
                //////                    Console.WriteLine("Processing... " + entAsset.GetAttributeValue<string>("hil_name") + " srno# " + srno);
                //////                    string fetchInvoices = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //////                        <entity name='hil_amcstaging'>
                //////                        <attribute name='hil_amcstagingid' /><attribute name='hil_warrantyenddate' /><attribute name='hil_warrantystartdate' />
                //////                        <order attribute='createdon' descending='false' />
                //////                        <filter type='and'>
                //////                            <condition attribute='statecode' operator='eq' value='0' />
                //////                            <condition attribute='hil_name' operator='eq' value='{_invNo}' />
                //////                            <condition attribute='hil_serailnumber' operator='eq' value='{srno}' />
                //////                        </filter>
                //////                        </entity>
                //////                    </fetch>";
                //////                    invcntColl = _service.RetrieveMultiple(new FetchExpression(fetchInvoices));
                //////                    int i = 0;
                //////                    foreach (Entity entcnt in invcntColl.Entities)
                //////                    {
                //////                        if (entcnt.Attributes.Contains("hil_warrantyenddate") && entcnt.Attributes.Contains("hil_warrantystartdate")) { i++; }
                //////                        if (i == 1) { continue; }
                //////                        else
                //////                        {
                //////                            _service.Update(new Entity()
                //////                            {
                //////                                LogicalName = entcnt.LogicalName,
                //////                                Id = entcnt.Id,
                //////                                Attributes = new AttributeCollection() { new System.Collections.Generic.KeyValuePair<string, object>("statecode", new OptionSetValue(1)), new System.Collections.Generic.KeyValuePair<string, object>("statuscode", new OptionSetValue(2)) }
                //////                            });
                //////                        }
                //////                        Console.WriteLine("Processing... " + i.ToString() + " || " + entAsset.GetAttributeValue<string>("hil_name") + " Invoice# " + entAsset.GetAttributeValue<string>("hil_name"));
                //////                    }
                //////                }
                //////                else
                //////                {
                //////                    Console.WriteLine("Processing... " + entAsset.GetAttributeValue<string>("hil_name") + "");
                //////                }
                //////            }

                //////            _service.Update(new Entity()
                //////            {
                //////                LogicalName = entAsset.LogicalName,
                //////                Id = entAsset.Id,
                //////                Attributes = new AttributeCollection() { new System.Collections.Generic.KeyValuePair<string, object>("statecode", new OptionSetValue(1)), new System.Collections.Generic.KeyValuePair<string, object>("statuscode", new OptionSetValue(2)) }
                //////            });
                //////        }
                //////    }
                //////    catch (Exception ex)
                //////    {
                //////        Console.WriteLine("ERROR... " + ex.Message);
                //////    }
                //////}
                ////#endregion

                ////#region Delete Duplicate Unit Warranty Lines
                ////string fetchDistinctAsset = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                ////    <entity name='hil_refreshjobs'>
                ////    <attribute name='hil_name' />
                ////    <attribute name='hil_channelpartnercode' />
                ////    <filter type='and'>
                ////        <condition attribute='statecode' operator='eq' value='0' />
                ////    </filter>
                ////    </entity>
                ////    </fetch>";
                ////while (true)
                ////{
                ////    try
                ////    {
                ////        EntityCollection custAssetColl = _service.RetrieveMultiple(new FetchExpression(fetchDistinctAsset));
                ////        if (custAssetColl.Entities.Count == 0) { break; }
                ////        foreach (Entity entAsset in custAssetColl.Entities)
                ////        {
                ////            string _SRNo = entAsset.GetAttributeValue<string>("hil_name");
                ////            string _WrtDate = entAsset.GetAttributeValue<string>("hil_channelpartnercode");
                ////            Console.WriteLine("Processing... " + entAsset.GetAttributeValue<string>("hil_name"));
                ////            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                ////              <entity name='hil_unitwarranty'>
                ////                <attribute name='hil_name' />
                ////                <order attribute='hil_warrantyenddate' descending='false' />
                ////                <filter type='and'>
                ////                  <condition attribute='hil_customerasset' operator='eq' value='{_SRNo}' />
                ////                  <condition attribute='hil_warrantyenddate' operator='eq' value='{_WrtDate}' />
                ////                  <condition attribute='statecode' operator='eq' value='0' />
                ////                </filter>
                ////                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' link-type='inner' alias='WT'>
                ////                  <attribute name='hil_type' />
                ////                  <filter type='and'>
                ////                    <condition attribute='hil_type' operator='eq' value='3' />
                ////                  </filter>
                ////                </link-entity>
                ////              </entity>
                ////            </fetch>";
                ////            int i = 0;
                ////            EntityCollection invColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                ////            foreach (Entity ent in invColl.Entities)
                ////            {
                ////                if (i++ == 0) { continue; }
                ////                else
                ////                {
                ////                    _service.Update(new Entity()
                ////                    {
                ////                        LogicalName = ent.LogicalName,
                ////                        Id = ent.Id,
                ////                        Attributes = new AttributeCollection() { new System.Collections.Generic.KeyValuePair<string, object>("statecode", new OptionSetValue(1)), new System.Collections.Generic.KeyValuePair<string, object>("statuscode", new OptionSetValue(2)) }
                ////                    });
                ////                }
                ////                Console.WriteLine("Processing... " + i.ToString() + " || " + entAsset.GetAttributeValue<string>("hil_name") + " Invoice# " + entAsset.GetAttributeValue<string>("hil_name"));
                ////            }
                ////            _service.Update(new Entity()
                ////            {
                ////                LogicalName = entAsset.LogicalName,
                ////                Id = entAsset.Id,
                ////                Attributes = new AttributeCollection() { new System.Collections.Generic.KeyValuePair<string, object>("statecode", new OptionSetValue(1)), new System.Collections.Generic.KeyValuePair<string, object>("statuscode", new OptionSetValue(2)) }
                ////            });
                ////        }

                ////    }
                ////    catch (Exception ex)
                ////    {
                ////        Console.WriteLine("ERROR... " + ex.Message);
                ////    }
                ////}
                ////#endregion
                ////IoTServiceCallRegistration _retVal = IoTCreateServiceCallDealerPortal(_service, new IoTServiceCallRegistration()
                ////{
                ////    AddressGuid= new Guid("ccff2a76-cd2a-ee11-bdf4-6045bda55a4d"),
                ////    ChiefComplaint= "fan not working properly",
                ////    CustomerGuid= new Guid("ca5ccf5a-cd2a-ee11-bdf4-6045bdac57b5"),
                ////    CustomerMobleNo= "9416881445",
                ////    NOCGuid= new Guid("28686003-dd14-e911-a94e-000d3af06091"),
                ////    ProductCategoryGuid=new Guid("32e44a8b-16fa-e811-a94d-000d3af06cd4"),
                ////    ProductSubCategoryGuid=new Guid("e2197022-410b-e911-a94f-000d3af00f43"),
                ////    SourceOfJob=6,
                ////    DealerCode= "CJK0036",
                ////    DealerName= "J.K. Electricals"

                ////});
                ////#region Publish Products
                ////EntityCollection entCol = null;
                ////FilterExpression criteria = new FilterExpression(LogicalOperator.And);
                ////criteria.AddCondition("statecode", ConditionOperator.NotEqual, 0);
                ////criteria.AddCondition("hil_hierarchylevel", ConditionOperator.NotEqual, 910590001);

                ////LinkEntity product = new LinkEntity
                ////{
                ////    Columns = new ColumnSet(false),
                ////    LinkFromEntityName = "hil_servicebom",
                ////    LinkToEntityName = "product",
                ////    LinkFromAttributeName = "hil_product",
                ////    LinkToAttributeName = "productid",
                ////    EntityAlias = "prd",
                ////    JoinOperator = JoinOperator.Inner,
                ////    LinkCriteria = criteria
                ////};
                ////QueryExpression queryExp = new QueryExpression("hil_servicebom");
                ////queryExp.ColumnSet = new ColumnSet("hil_product");
                ////queryExp.Distinct = true;
                ////queryExp.PageInfo = new PagingInfo
                ////{
                ////    PageNumber = 1,
                ////    Count = 5000
                ////};
                ////queryExp.LinkEntities.Add(product);

                ////string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                ////    <entity name='msdyn_workorderproduct'>
                ////    <attribute name='msdyn_product' />
                ////    <order attribute='msdyn_product' descending='false' />
                ////    <filter type='and'>
                ////        <condition attribute='hil_markused' operator='eq' value='1' />
                ////    </filter>
                ////    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ak'>
                ////        <filter type='and'>
                ////        <condition attribute='msdyn_substatus' operator='eq' uiname='Closed' uitype='msdyn_workordersubstatus' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                ////        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='2023-07-01' />
                ////        <condition attribute='hil_owneraccount' operator='eq' uiname='NCR Cooling Care' uitype='account' value='{330FB73A-310B-E911-A94D-000D3AF06C56}' />
                ////        </filter>
                ////    </link-entity>
                ////    <link-entity name='product' from='productid' to='msdyn_product' link-type='inner' alias='al'>
                ////        <filter type='and'>
                ////        <condition attribute='hil_hierarchylevel' operator='eq' value='5' />
                ////        </filter>
                ////    </link-entity>
                ////    </entity>
                ////    </fetch>";

                ////int row = 1;
                ////int totalRowCount = 0;

                ////while (true)
                ////{
                ////    entCol = _service.RetrieveMultiple(queryExp);
                ////    //entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                ////    totalRowCount += entCol.Entities.Count;

                ////    foreach (Entity ent in entCol.Entities)
                ////    {
                ////        Console.WriteLine("Processing..." + ent.GetAttributeValue<EntityReference>("hil_product").Name);
                ////        //Console.WriteLine("Processing..." + ent.GetAttributeValue<EntityReference>("msdyn_product").Name);
                ////        try
                ////        {
                ////            _service.Execute(new SetStateRequest
                ////            {
                ////                EntityMoniker = ent.GetAttributeValue<EntityReference>("hil_product"),
                ////                State = new OptionSetValue((int)ProductState.Active),
                ////                Status = new OptionSetValue(1)
                ////            });
                ////            Entity entTemp = new Entity("product", ent.GetAttributeValue<EntityReference>("hil_product").Id);
                ////            entTemp["msdyn_fieldserviceproducttype"] = new OptionSetValue(690970000);
                ////            _service.Update(entTemp);
                ////        }
                ////        catch (Exception ex)
                ////        {
                ////            Console.WriteLine("Error while Processing..." + ent.GetAttributeValue<EntityReference>("hil_product").Name + ":: " + ex.Message);
                ////        }
                ////        Console.WriteLine("Processing..." + row++.ToString() + "/" + totalRowCount.ToString() + ":: " + ent.GetAttributeValue<EntityReference>("hil_product").Name);
                ////    }
                ////    if (!entCol.MoreRecords) break;
                ////    queryExp.PageInfo.PageNumber++;
                ////    queryExp.PageInfo.PagingCookie = entCol.PagingCookie;
                ////}
                ////Console.WriteLine("TOTAL LINES: " + totalRowCount.ToString());
                ////#endregion

                //#region Call Masking Airtel IQ
                ////CallMasking _callmasking = new CallMasking();
                ////ResposeDataCallMasking _reData = _callmasking.GetCustomerOpenJobs(_service, new RequestDataCallMasking()
                ////{
                ////    MobileNumber = "7678373445",
                ////    JobNumber = ""
                ////});
                ////Console.WriteLine(_reData.JobsData.Count);
                //#endregion
                ////string _fetchXML = $@"<fetch>
                ////<entity name='userquery'>
                ////<!-- Columns -->
                ////<attribute name='userqueryid' />
                ////<attribute name='fetchxml' />
                ////<attribute name='name' />
                ////<!-- Filter By -->
                ////<filter type='and'>
                //// <condition attribute='name' operator='eq' value='SAW to be rejected at Closing 1' />
                ////</filter>
                ////</entity>
                ////</fetch>";
                ////EntityCollection jobsColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                ////if (jobsColl.Entities.Count > 0)
                ////{
                ////    Console.WriteLine(jobsColl.Entities[0].GetAttributeValue<string>("fetchxml"));
                ////}
                ////int _pageNum = 3;
                ////int _pageSize = 10;

                ////string _fetchXML = $@"<fetch version='1.0' count='{_pageSize}' page='{_pageNum}'>
                ////<entity name='hil_city'>
                ////<attribute name='hil_name' />
                ////<attribute name='modifiedon' />
                ////<order attribute='modifiedon' descending='false' />
                ////</entity>
                ////</fetch>";
                ////EntityCollection jobsColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                ////if (jobsColl.Entities.Count > 0)
                ////{
                ////    Console.WriteLine("Page # " + _pageNum.ToString());
                ////    foreach (Entity ent in jobsColl.Entities)
                ////    {
                ////        Console.WriteLine(ent.GetAttributeValue<string>("hil_name") + "|" + ent.GetAttributeValue<DateTime>("modifiedon").ToString());
                ////    }
                ////}
                ////CallMasking _obj = new CallMasking();
                ////ResposeDataCallMasking _retVal = _obj.GetCustomerOpenJobs(new RequestDataCallMasking()
                ////{
                ////    MobileNumber = "8285906486"
                ////}, _service);
                ////Console.WriteLine(_retVal.JobsData.Count.ToString());

                ////CallMasking _obj = new CallMasking();
                ////CDR_Response _retVal = _obj.PushCDR(new CDR_Request()
                ////{
                ////    Correlation_ID = "Xchange4ddc3e61-bb73-415c-ab3c-bc9170669008",
                ////    Caller_Id = "8765456789",
                ////    Caller_Name = "Deepanshu",
                ////    Caller_Number = "8810570149",
                ////    Call_Type = "OutBound",
                ////    Overall_Call_Duration = "00:58",
                ////    Time = "08:42:06",
                ////    Date = "14/07/2020",
                ////    Overall_Call_Status = "Answered",
                ////    Recording = "https://openapi.airtel.in/gateway/airtel-xchange/ironman-data-transfer/download/recordingFile?token=gLTIM5uvaO6TqDKclWycfdjdC+z6poFo9UjzcnTB3VJnV5VPXmeaBWCurwyrCD/arWS+b+M/C23OW9cr2uu+dY7Hmls+8Y2hWF6DFD0VZMc=",
                ////    Destination_Number = "986546782",
                ////    Destination_Name ="Azad Kumar"
                ////}, _service);
                ////Console.WriteLine(_retVal.ResultStatus);
                ////Execute(_service);
            }
                }

        private static void SendEmailOnActivityComplete(Entity CaseGrievanceAct, IOrganizationService service)
        {
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_grievancehandlingactivity'>
                    <attribute name='regardingobjectid' />
                    <attribute name='ownerid' />
                    <order attribute='subject' descending='false' />
                    <filter type='and'>
                      <condition attribute='activityid' operator='eq' value='{CaseGrievanceAct.Id}' />
                    </filter>
                    <link-entity name='hil_caseassignmentmatrixline' from='hil_caseassignmentmatrixlineid' to='hil_caseassignmentmatrixlineid' visible='false' link-type='outer' alias='amt'>
                      <attribute name='hil_level' />
                      <attribute name='hil_caseassignmentmatrixid' />
                    </link-entity>
                    <link-entity name='incident' from='incidentid' to='regardingobjectid' visible='false' link-type='outer' alias='case'>
                      <attribute name='title' />
                      <attribute name='ticketnumber' />
                    </link-entity>
                  </entity>
                </fetch>";
            EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (_entCol.Entities.Count > 0)
            {
                int Level = Convert.ToInt32(_entCol.Entities[0].GetAttributeValue<AliasedValue>("amt.hil_level").Value);
                Guid CaseAssignmentMatrixId = ((EntityReference)_entCol.Entities[0].GetAttributeValue<AliasedValue>("amt.hil_caseassignmentmatrixid").Value).Id;
                EntityReference _entCase = _entCol.Entities[0].GetAttributeValue<EntityReference>("regardingobjectid");
                EntityReference _entActivityOwner = _entCol.Entities[0].GetAttributeValue<EntityReference>("ownerid");
                string CaseNumber = _entCol.Entities[0].Contains("case.ticketnumber") ? _entCol.Entities[0].GetAttributeValue<AliasedValue>("case.ticketnumber").Value.ToString() : null;
                string caseTitle = _entCol.Entities[0].Contains("case.title") ? _entCol.Entities[0].GetAttributeValue<AliasedValue>("case.title").Value.ToString() : null;
                string recordURL = $"https://havells.crm8.dynamics.com/main.aspx?appid=668eb624-0610-e911-a94e-000d3af06a98&forceUCI=1&pagetype=entityrecord&etn=incident&id={_entCase.Id}";

                QueryExpression query = new QueryExpression("hil_caseassignmentmatrixline");
                query.ColumnSet = new ColumnSet("hil_assigneeuser");
                query.Criteria.Filters.Add(new FilterExpression(LogicalOperator.And));
                query.Criteria.AddCondition("hil_level", ConditionOperator.Equal, (Level + 1));
                query.Criteria.AddCondition("hil_caseassignmentmatrixid", ConditionOperator.Equal, CaseAssignmentMatrixId);

                EntityCollection coll = service.RetrieveMultiple(query);
                if (coll.Entities.Count > 0)
                {
                    EntityReference EscUser = coll.Entities[0].GetAttributeValue<EntityReference>("hil_assigneeuser");

                    Entity fromActivityParty = new Entity("activityparty");
                    Entity toActivityParty = new Entity("activityparty");

                    fromActivityParty["partyid"] = new EntityReference("queue", new Guid("9b0480a8-e30f-e911-a94e-000d3af06a98"));
                    toActivityParty["partyid"] = EscUser;

                    Entity email = new Entity("email");
                    email["from"] = new Entity[] { fromActivityParty };
                    email["to"] = new Entity[] { toActivityParty };
                    email["regardingobjectid"] = _entCase;
                    email["subject"] = $"Activity mark completed for case nubmer {CaseNumber}";
                    email["description"] = $"Dear Team, <br/><br/>Activity mark completed by {_entActivityOwner.Name} for the case with case-ID {CaseNumber} regarding {caseTitle}. <br/><br/> To open the case <a href={recordURL}>Click Here</a> <br/> <br/> Regards Team CRM";
                    email["directioncode"] = true;
                    Guid emailId = service.Create(email);
                    //Use the SendEmail message to send an e-mail message.
                    SendEmailRequest sendEmailRequest = new SendEmailRequest
                    {
                        EmailId = emailId,
                        TrackingToken = "",
                        IssueSend = true
                    };
                    SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);
                }
            }
        }
        private static void HavellsOneWebsiteAMCPlanData(IOrganizationService _service)
        {
            try
            {
                int _pageNum = 1;
                while (true)
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' count='5000' page='{_pageNum}'>
                    <entity name='hil_servicebom'>
                    <attribute name='hil_product' />
                    <attribute name='hil_servicebomid' />
                    <attribute name='hil_productcategory' />
                    <order attribute='hil_product' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='at'>
                        <filter type='and'>
                        <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                        </filter>
                    </link-entity>
                    <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='am'>
                        <attribute name='description' />
                        <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='au'>
                        <filter type='and'>
                            <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                        </filter>
                        </link-entity>
                    </link-entity>
                    </entity>
                </fetch>";
                    EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (_entCol.Entities.Count > 0)
                    {
                        Entity _entAMCPlan = null;
                        EntityCollection _entColPlan = null;
                        int _rowCount = 1;
                        foreach (Entity ent in _entCol.Entities)
                        {
                            Console.WriteLine($"Processing... {_rowCount++.ToString()}/{_entCol.Entities.Count}");

                            _entAMCPlan = new Entity("hil_havellsonewebsiteamcplans");

                            _entAMCPlan["hil_planname"] = ent.Contains("hil_product") ? ent.GetAttributeValue<EntityReference>("hil_product").Name : null;
                            _entAMCPlan["hil_planid"] = ent.Contains("hil_product") ? ent.GetAttributeValue<EntityReference>("hil_product").Id.ToString() : null;

                            if (ent.Contains("hil_productcategory"))
                            {
                                _entAMCPlan["hil_name"] = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                                _entAMCPlan["hil_modelname"] = _service.Retrieve("product", ent.GetAttributeValue<EntityReference>("hil_productcategory").Id, new ColumnSet("description")).GetAttributeValue<string>("description");
                                _entAMCPlan["hil_modelid"] = ent.GetAttributeValue<EntityReference>("hil_productcategory").Id.ToString();
                            }

                            string _fetchXMLAMC = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_productcatalog'>
                            <attribute name='hil_plantclink' />
                            <attribute name='hil_planperiod' />
                            <attribute name='hil_notcovered' />
                            <attribute name='hil_coverage' />
                            <attribute name='hil_amctandc' />
                            <filter type='and'>
                                <condition attribute='hil_productcode' operator='eq' value='{ent.GetAttributeValue<EntityReference>("hil_product").Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                        </fetch>";
                            _entColPlan = _service.RetrieveMultiple(new FetchExpression(_fetchXMLAMC));
                            if (_entColPlan.Entities.Count > 0)
                            {
                                _entAMCPlan["hil_plantcurl"] = _entColPlan.Entities[0].Contains("hil_plantclink") ? _entColPlan.Entities[0].GetAttributeValue<string>("hil_plantclink") : null;
                                _entAMCPlan["hil_planperiod"] = _entColPlan.Entities[0].Contains("hil_planperiod") ? _entColPlan.Entities[0].GetAttributeValue<string>("hil_planperiod") : null;
                                _entAMCPlan["hil_noncoverage"] = _entColPlan.Entities[0].Contains("hil_notcovered") ? _entColPlan.Entities[0].GetAttributeValue<string>("hil_notcovered") : null;
                                _entAMCPlan["hil_mrp"] = _service.Retrieve("product", ent.GetAttributeValue<EntityReference>("hil_product").Id, new ColumnSet("hil_amount")).GetAttributeValue<Money>("hil_amount");
                                _entAMCPlan["hil_coverage"] = _entColPlan.Entities[0].Contains("hil_coverage") ? _entColPlan.Entities[0].GetAttributeValue<string>("hil_coverage") : null;
                            }
                            _service.Create(_entAMCPlan);
                        }
                    }
                    _pageNum += 1;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private static void CalculateRMCost(IOrganizationService _service, Guid _rmCostSheetId)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                          <entity name='hil_rmcostsheetline'>
                            <attribute name='hil_rmcostsheet' groupby='true' alias='header'/>
                            <attribute name='hil_cost' aggregate='sum' alias='totalCost'/>
                            <filter type='and'>
                              <condition attribute='hil_rmcostsheet' operator='eq' value='{_rmCostSheetId}' />
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                          </entity>
                        </fetch>";
                EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_entCol.Entities.Count > 0) {
                    Money _totalCost = (Money)(_entCol.Entities[0].GetAttributeValue<AliasedValue>("totalCost").Value);

                    Entity _entRMCostSheet = new Entity("hil_rmcostsheet", _rmCostSheetId);
                    _entRMCostSheet["hil_rmcost"] = _totalCost;
                    _service.Update(_entRMCostSheet);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private static void AssignCase(IOrganizationService _service, Entity _entity, EntityReference _assignee, EntityReference _assignmentMatrix)
        {
            try
            {
                Entity _entCase = new Entity(_entity.LogicalName, _entity.Id);
                _entCase["hil_assignmentmatrix"] = _assignmentMatrix;
                _entCase["ownerid"] = _assignee;
                _service.Update(_entCase);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        static void UpdateJobWarrantyDetails(Guid _jobGuId, IOrganizationService _service)
        {
            DateTime _jobCreatedOn = new DateTime(1900, 1, 1);
            DateTime _assetPurchaseDate = new DateTime(1900, 1, 1);
            DateTime _unitWarrStartDate = new DateTime(1900, 1, 1);
            QueryExpression qrExp;
            EntityCollection entCol;
            EntityReference _unitWarranty = null;
            Guid _warrantyTemplateId = Guid.Empty;
            bool laborInWarranty = false;
            int _jobWarrantyStatus = 2; //OutWarranty
            int _jobWarrantySubStatus = 0;
            int _warrantyTempType = 0;
            double _jobMonth = 0;
            bool SparePartUsed = false;
            OptionSetValue _opChannelPartnerCategory = null;

            Entity entJob = _service.Retrieve(msdyn_workorder.EntityLogicalName, _jobGuId, new ColumnSet("hil_owneraccount", "createdon", "msdyn_customerasset"));
            if (entJob != null)
            {
                if (entJob.Attributes.Contains("hil_owneraccount"))
                {
                    Entity entTemp = _service.Retrieve(Account.EntityLogicalName, entJob.GetAttributeValue<EntityReference>("hil_owneraccount").Id, new ColumnSet("hil_category"));
                    if (entTemp != null)
                    {
                        if (entTemp.Attributes.Contains("hil_category"))
                        {
                            _opChannelPartnerCategory = entTemp.GetAttributeValue<OptionSetValue>("hil_category");
                        }
                    }
                }
                _jobCreatedOn = entJob.GetAttributeValue<DateTime>("createdon");
                if (entJob.Attributes.Contains("msdyn_customerasset"))
                {
                    Entity entCustAsset = _service.Retrieve(msdyn_customerasset.EntityLogicalName, entJob.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_invoicedate"));
                    if (entCustAsset != null)
                    {
                        if (entCustAsset.Attributes.Contains("hil_invoicedate"))
                        {
                            _assetPurchaseDate = entCustAsset.GetAttributeValue<DateTime>("hil_invoicedate");
                        }
                    }

                    qrExp = new QueryExpression("msdyn_workorderproduct");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderproductid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_markused", ConditionOperator.Equal, true);
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count > 0) { SparePartUsed = true; }

                    qrExp = new QueryExpression("msdyn_workorderincident");
                    qrExp.ColumnSet = new ColumnSet("msdyn_workorderincidentid");
                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                    qrExp.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, _jobGuId);
                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    qrExp.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, 3); // Warranty Void
                    entCol = _service.RetrieveMultiple(qrExp);
                    if (entCol.Entities.Count == 0)
                    {
                        qrExp = new QueryExpression("hil_unitwarranty");
                        qrExp.ColumnSet = new ColumnSet("hil_warrantytemplate", "hil_warrantystartdate", "hil_warrantyenddate");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, entJob.GetAttributeValue<EntityReference>("msdyn_customerasset").Id);
                        //qrExp.Criteria.AddCondition("hil_warrantystartdate", ConditionOperator.OnOrBefore, new DateTime(_jobCreatedOn.Year, _jobCreatedOn.Month, _jobCreatedOn.Day));
                        //qrExp.Criteria.AddCondition("hil_warrantyenddate", ConditionOperator.OnOrAfter, new DateTime(_jobCreatedOn.Year, _jobCreatedOn.Month, _jobCreatedOn.Day));
                        qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entCol = _service.RetrieveMultiple(qrExp);
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity Wt in entCol.Entities)
                            {
                                DateTime iValidTo = (DateTime)Wt["hil_warrantyenddate"];
                                DateTime iValidFrom = (DateTime)Wt["hil_warrantystartdate"];
                                if (_jobCreatedOn >= iValidFrom && _jobCreatedOn <= iValidTo)
                                {
                                    _jobWarrantyStatus = 1;
                                    _unitWarranty = Wt.ToEntityReference();
                                    _warrantyTemplateId = Wt.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;
                                    Entity _entTemp = _service.Retrieve(hil_warrantytemplate.EntityLogicalName, _warrantyTemplateId, new ColumnSet("hil_type"));
                                    if (_entTemp != null)
                                    {
                                        _warrantyTempType = _entTemp.GetAttributeValue<OptionSetValue>("hil_type").Value;
                                        if (_warrantyTempType == 1) { _jobWarrantySubStatus = 1; }
                                        else if (_warrantyTempType == 2) { _jobWarrantySubStatus = 2; }
                                        else if (_warrantyTempType == 7) { _jobWarrantySubStatus = 3; }
                                        else if (_warrantyTempType == 3) { _jobWarrantySubStatus = 4; }
                                    }
                                    _unitWarrStartDate = Wt.GetAttributeValue<DateTime>("hil_warrantystartdate");
                                    TimeSpan difference = (_jobCreatedOn - _unitWarrStartDate);
                                    _jobMonth = Math.Round((difference.Days * 1.0 / 30.42), 0);
                                    qrExp = new QueryExpression("hil_labor");
                                    qrExp.ColumnSet = new ColumnSet("hil_laborid", "hil_includedinwarranty", "hil_validtomonths", "hil_validfrommonths");
                                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    qrExp.Criteria.AddCondition("hil_warrantytemplateid", ConditionOperator.Equal, _warrantyTemplateId);
                                    qrExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                    entCol = _service.RetrieveMultiple(qrExp);
                                    if (entCol.Entities.Count == 0) { laborInWarranty = true; }
                                    else
                                    {
                                        if (_jobMonth >= entCol.Entities[0].GetAttributeValue<int>("hil_validfrommonths") && _jobMonth <= entCol.Entities[0].GetAttributeValue<int>("hil_validtomonths"))
                                        {
                                            OptionSetValue _laborType = entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                            laborInWarranty = _laborType.Value == 1 ? true : false;
                                        }
                                        else
                                        {
                                            OptionSetValue _laborType = entCol.Entities[0].GetAttributeValue<OptionSetValue>("hil_includedinwarranty");
                                            laborInWarranty = !(_laborType.Value == 1 ? true : false);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        _jobWarrantyStatus = 3;
                    }
                }
            }

            Entity entJobUpdate = new Entity(msdyn_workorder.EntityLogicalName, _jobGuId);
            entJobUpdate["hil_warrantystatus"] = new OptionSetValue(_jobWarrantyStatus);
            if (_jobWarrantyStatus == 1)
            {
                entJobUpdate["hil_warrantysubstatus"] = new OptionSetValue(_jobWarrantySubStatus);
            }
            if (_unitWarranty != null)
            {
                entJobUpdate["hil_unitwarranty"] = _unitWarranty;
            }
            entJobUpdate["hil_laborinwarranty"] = laborInWarranty;
            if (_assetPurchaseDate.Year != 1900)
            {
                entJobUpdate["hil_purchasedate"] = _assetPurchaseDate;
            }
            entJobUpdate["hil_sparepartuse"] = SparePartUsed;
            if (_opChannelPartnerCategory != null)
            {
                entJobUpdate["hil_channelpartnercategory"] = _opChannelPartnerCategory;
            }
            _service.Update(entJobUpdate);
        }
        static void UpdateDuplicateLeads(IOrganizationService service) { 
        
        }
        public static void Execute(IOrganizationService service)
        {
            Guid jobId = Guid.Empty;
            bool inputsValidated = true;
            string strInputsValidationsummary = string.Empty;
            string addressRemarks = string.Empty;

            try
            {
                DateTime? hil_expecteddeliverydate = null;

                #region Getting Data from Excel Sheet Columns
                string hil_addressline1 = "Flat no 45,2nd Flr,Building No 48";

                string hil_addressline2 = "Anarya chs,Jasmine Mill Road";

                string hil_alternatenumber = hil_alternatenumber = "";

                string hil_area = "";

                OptionSetValue hil_callertype = new OptionSetValue(1);
                string hil_callsubtype = "Installation";
                string hil_customerfirstname = "Abdul";
                string hil_customerlastname = "Jabbar";
                string hil_customermobileno = "9967654517";
                string hil_dealercode = "Croma";

                hil_expecteddeliverydate = Convert.ToDateTime("13-Jul-2023");

                string hil_landmark = "";

                string hil_natureofcomplaint = "Installation";
                string hil_pincode = "400019";
                string hil_productsubcategory = "STORAGE WATER HEATER HAVELLS";
                string hil_salutation = "Mr";
                #endregion

                #region Validating Excel file data
                //if (hil_salutation == string.Empty || hil_salutation == null) { strInputsValidationsummary = "Salutation is required."; inputsValidated = false; }
                if (hil_customermobileno == string.Empty || hil_customermobileno == null) { strInputsValidationsummary += "\n Customer Mobile Number is required."; inputsValidated = false; }
                if (hil_customerfirstname == string.Empty || hil_customerfirstname == null) { strInputsValidationsummary += "\n Customer First Name is required."; inputsValidated = false; }
                if (hil_addressline1 == string.Empty || hil_addressline1 == null) { strInputsValidationsummary += "\n Address Line 1 is required."; inputsValidated = false; }
                if (hil_pincode == string.Empty || hil_pincode == null) { strInputsValidationsummary += "\n PIN Code is required."; inputsValidated = false; }
                //if (hil_area == string.Empty || hil_area == null) { strInputsValidationsummary += "\n Area is required."; inputsValidated = false; }
                if (hil_productsubcategory == string.Empty || hil_productsubcategory == null) { strInputsValidationsummary += "\n Product Sub Category is required."; inputsValidated = false; }
                if (hil_natureofcomplaint == string.Empty || hil_natureofcomplaint == null) { strInputsValidationsummary += "\n Nature of Complaint is required."; inputsValidated = false; }
                if (hil_callsubtype == string.Empty || hil_callsubtype == null) { strInputsValidationsummary += "\n Call SubType is required."; inputsValidated = false; }
                if (hil_callertype == null) { strInputsValidationsummary += "\n Caller Type is required."; inputsValidated = false; }

                if (hil_callsubtype.ToUpper() == "INSTALLATION")
                {
                    if (hil_expecteddeliverydate == null) { hil_expecteddeliverydate = DateTime.Now; }
                }
                #endregion

                #region Declaring Local Variables
                OptionSetValue salutation;
                Entity entTemp;
                Entity entConsumer = null;

                QueryExpression queryExp;
                EntityCollection entCol;
                EntityReference erArea = null;
                EntityReference erPinCode = null;
                EntityReference erBizGeoMapping = null;
                EntityReference erProductSubCategoryStagging = null;
                EntityReference erproductSubCategory = null;
                EntityReference erproductCategory = null;
                EntityReference erCallSubType = null;
                EntityReference erAddress = null;
                EntityReference erNatureOfComplaint = null;
                EntityReference erConsumerCategory = null;
                EntityReference erConsumerType = null;
                Guid _customerGuid = Guid.Empty;
                bool _duplicateWO = false;
                string _duplicateWOIds = string.Empty;
                Entity entity = new Entity();
                #endregion

                #region Creating Consumer, Address, and Work Order Records
                if (inputsValidated)
                {

                    entTemp = ExecuteScalar(service, "hil_consumercategory", "hil_name", "End User", new string[] { "hil_consumercategoryid" }, ref entity);
                    if (entTemp != null)
                    { erConsumerCategory = entTemp.ToEntityReference(); }
                    entTemp = ExecuteScalar(service, "hil_consumertype", "hil_name", "B2C", new string[] { "hil_consumertypeid" }, ref entity);
                    if (entTemp != null)
                    { erConsumerType = entTemp.ToEntityReference(); }
                    if (hil_salutation.ToUpper() == "MR." || hil_salutation.ToUpper() == "MR") { salutation = new OptionSetValue(2); }
                    else if (hil_salutation.ToUpper() == "MISS." || hil_salutation.ToUpper() == "MISS") { salutation = new OptionSetValue(1); }
                    else if (hil_salutation.ToUpper() == "MRS." || hil_salutation.ToUpper() == "MRS") { salutation = new OptionSetValue(3); }
                    else if (hil_salutation.ToUpper() == "DR." || hil_salutation.ToUpper() == "DR") { salutation = new OptionSetValue(4); }
                    else if (hil_salutation.ToUpper() == "M/S." || hil_salutation.ToUpper() == "M/S") { salutation = new OptionSetValue(5); }
                    else { salutation = new OptionSetValue(2); }

                    queryExp = new QueryExpression("contact");
                    queryExp.ColumnSet = new ColumnSet("fullname", "emailaddress1", "contactid");
                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                    queryExp.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, hil_customermobileno));
                    entCol = service.RetrieveMultiple(queryExp);

                    if (entCol.Entities.Count == 1)
                    {
                        _customerGuid = entCol.Entities[0].Id;
                        entTemp = ExecuteScalar(service, "hil_stagingdivisonmaterialgroupmapping", "hil_name", hil_productsubcategory, new string[] { "hil_productcategorydivision", "hil_productsubcategorymg" }, ref entity);
                        if (entTemp != null)
                        {
                            erproductCategory = entTemp.GetAttributeValue<EntityReference>("hil_productcategorydivision");
                            #region Duplicate Work Order Check {Same Customer/Prod Category and Open Job in last 30 days}
                            queryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            queryExp.ColumnSet = new ColumnSet("msdyn_name");
                            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, _customerGuid));
                            queryExp.Criteria.AddCondition(new ConditionExpression("hil_productcategory", ConditionOperator.Equal, erproductCategory.Id));
                            queryExp.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.LastXDays, 30));
                            queryExp.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.NotIn, new object[] { new Guid("1527FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("6C8F2123-5106-EA11-A811-000D3AF057DD"), new Guid("2927FA6C-FA0F-E911-A94E-000D3AF060A1"), new Guid("7E85074C-9C54-E911-A951-000D3AF0677F") }));
                            queryExp.AddOrder("createdon", OrderType.Descending);
                            entCol = service.RetrieveMultiple(queryExp);
                            if (entCol.Entities.Count > 0)
                            {
                                _duplicateWO = true;
                                foreach (Entity ent in entCol.Entities)
                                {
                                    _duplicateWOIds += ent.GetAttributeValue<string>("msdyn_name") + ",";
                                }
                            }
                            #endregion
                        }
                    }
                    else
                    {
                        string[] fullName = hil_customerfirstname.Split(' ');
                        if (fullName.Length > 1)
                        {
                            hil_customerfirstname = fullName[0];
                            hil_customerlastname = fullName[1];
                        }

                        try
                        {
                            Contact cnt = new Contact()
                            {
                                FirstName = hil_customerfirstname,
                                LastName = hil_customerlastname,
                                MobilePhone = hil_customermobileno,
                                hil_Salutation = salutation
                            };
                            _customerGuid = service.Create(cnt);
                        }
                        catch (Exception ex)
                        {
                            entity["hil_description"] = "Error While Creating Customer!!! " + ex.Message;
                            entity["hil_jobstatus"] = false;
                        }
                    }
                    if (!_duplicateWO)
                    {
                        if (_customerGuid != Guid.Empty)
                        {
                            entConsumer = service.Retrieve("contact", _customerGuid, new ColumnSet("fullname", "emailaddress1", "contactid"));
                            Entity entPincode = ExecuteScalar(service, "hil_pincode", "hil_name", hil_pincode, new string[] { "hil_pincodeid" }, ref entity);
                            if (entPincode != null)
                            {
                                erPinCode = entPincode.ToEntityReference();
                            }
                            erAddress = GetCustomerAddress(service, _customerGuid, erPinCode != null ? erPinCode.Id : Guid.Empty);

                            if (erAddress == null)
                            {
                                entTemp = ExecuteScalar(service, "hil_area", "hil_name", hil_area, new string[] { "hil_areaid" }, ref entity);
                                if (entTemp != null) { erArea = entTemp.ToEntityReference(); } else { addressRemarks = "Area does not exist in Area Master."; }
                                //entTemp = ExecuteScalar(service, "hil_pincode", "hil_name", hil_pincode, new string[] { "hil_pincodeid" }, ref entity);
                                if (entPincode != null)
                                {
                                    //erPinCode = entTemp.ToEntityReference();
                                    queryExp = new QueryExpression("hil_businessmapping");
                                    queryExp.ColumnSet = new ColumnSet("hil_businessmappingid");
                                    queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, erPinCode.Id));
                                    if (erArea != null)
                                    {
                                        queryExp.Criteria.AddCondition(new ConditionExpression("hil_area", ConditionOperator.Equal, erArea.Id));
                                    }
                                    queryExp.TopCount = 1;
                                    entCol = service.RetrieveMultiple(queryExp);
                                    if (entCol.Entities.Count >= 1)
                                    {
                                        erBizGeoMapping = entCol.Entities[0].ToEntityReference();
                                        try
                                        {
                                            hil_address addr = new hil_address()
                                            {
                                                hil_Street1 = hil_addressline1,
                                                hil_Street2 = hil_addressline2,
                                                hil_Street3 = hil_landmark,
                                                hil_AddressType = new OptionSetValue(1),
                                                hil_BusinessGeo = erBizGeoMapping,
                                                hil_Customer = new EntityReference("contact", entConsumer.Id)
                                            };
                                            Guid _addressGuid = service.Create(addr);
                                            erAddress = new EntityReference("hil_address", _addressGuid);
                                        }
                                        catch (Exception ex)
                                        {
                                            entity["hil_description"] = "Error While Creating Consumer's Address." + ex.Message;
                                            entity["hil_jobstatus"] = false;
                                        }
                                    }
                                    else
                                    {
                                        entity["hil_description"] = "No Business Geo Mapping found for PIN Code.";
                                        entity["hil_jobstatus"] = false;
                                    }
                                }
                                else
                                {
                                    entity["hil_description"] = "PIN Code does not exist.";
                                    entity["hil_jobstatus"] = false;
                                }
                            }
                        }

                        if (erAddress == null)
                        {
                            entTemp = ExecuteScalar(service, "hil_pincode", "hil_name", hil_pincode, new string[] { "hil_pincodeid" }, ref entity);
                            if (entTemp != null)
                            {
                                erPinCode = entTemp.ToEntityReference();
                                queryExp = new QueryExpression("hil_businessmapping");
                                queryExp.ColumnSet = new ColumnSet("hil_businessmappingid");
                                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                                queryExp.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, erPinCode.Id));
                                if (erArea != null)
                                {
                                    queryExp.Criteria.AddCondition(new ConditionExpression("hil_area", ConditionOperator.Equal, erArea.Id));
                                }
                                queryExp.TopCount = 1;
                                entCol = service.RetrieveMultiple(queryExp);
                                if (entCol.Entities.Count >= 1)
                                {
                                    erBizGeoMapping = entCol.Entities[0].ToEntityReference();
                                    try
                                    {
                                        hil_address addr = new hil_address()
                                        {
                                            hil_Street1 = hil_addressline1,
                                            hil_Street2 = hil_addressline2,
                                            hil_Street3 = hil_landmark,
                                            hil_AddressType = new OptionSetValue(1),
                                            hil_BusinessGeo = erBizGeoMapping,
                                            hil_Customer = new EntityReference("contact", entConsumer.Id)
                                        };
                                        Guid _addressGuid = service.Create(addr);
                                        erAddress = new EntityReference("hil_address", _addressGuid);
                                    }
                                    catch (Exception ex)
                                    {
                                        entity["hil_description"] = "Error While Creating Consumer's Address." + ex.Message;
                                        entity["hil_jobstatus"] = false;
                                    }
                                }
                                else
                                {
                                    entity["hil_description"] = "No Business Geo Mapping found for PIN Code.";
                                    entity["hil_jobstatus"] = false;
                                }
                            }
                        }
                        if (entConsumer != null && erAddress != null)
                        {
                            entTemp = ExecuteScalar(service, "hil_stagingdivisonmaterialgroupmapping", "hil_name", hil_productsubcategory, new string[] { "hil_productcategorydivision", "hil_productsubcategorymg" }, ref entity);
                            if (entTemp != null)
                            {
                                erProductSubCategoryStagging = entTemp.ToEntityReference();
                                erNatureOfComplaint = GetNatureOfComplaint(service, hil_natureofcomplaint, entTemp.GetAttributeValue<EntityReference>("hil_productsubcategorymg").Id);
                                erproductCategory = entTemp.GetAttributeValue<EntityReference>("hil_productcategorydivision");
                                erproductSubCategory = entTemp.GetAttributeValue<EntityReference>("hil_productsubcategorymg");
                                entTemp = ExecuteScalar(service, "hil_callsubtype", "hil_name", hil_callsubtype, new string[] { "hil_callsubtypeid" }, ref entity);
                                if (entTemp != null)
                                {
                                    erCallSubType = entTemp.ToEntityReference();
                                }
                                msdyn_workorder enPMSWorkorder = new msdyn_workorder();
                                enPMSWorkorder.hil_CustomerRef = entConsumer.ToEntityReference();
                                enPMSWorkorder.hil_customername = entConsumer.GetAttributeValue<string>("fullname");
                                enPMSWorkorder.hil_mobilenumber = hil_customermobileno;
                                enPMSWorkorder.hil_Alternate = hil_alternatenumber;
                                enPMSWorkorder.hil_Address = erAddress;
                                enPMSWorkorder.hil_Productcategory = erproductCategory;
                                enPMSWorkorder.hil_ProductCatSubCatMapping = erProductSubCategoryStagging;
                                if (erConsumerType != null)
                                {
                                    enPMSWorkorder["hil_consumertype"] = erConsumerType;
                                }
                                if (erConsumerCategory != null)
                                {
                                    enPMSWorkorder["hil_consumercategory"] = erConsumerCategory;
                                }
                                if (erNatureOfComplaint != null)
                                {
                                    enPMSWorkorder.hil_natureofcomplaint = erNatureOfComplaint;
                                }
                                enPMSWorkorder["hil_newserialnumber"] = hil_dealercode;
                                if (erCallSubType != null)
                                {
                                    enPMSWorkorder.hil_CallSubType = erCallSubType;
                                }
                                enPMSWorkorder.hil_quantity = 1;
                                enPMSWorkorder["hil_callertype"] = hil_callertype;
                                if (hil_expecteddeliverydate != null)
                                {
                                    enPMSWorkorder["hil_expecteddeliverydate"] = hil_expecteddeliverydate;
                                }
                                enPMSWorkorder.hil_SourceofJob = new OptionSetValue(14); // {SourceofJob:"Excel Upload"}
                                enPMSWorkorder.msdyn_ServiceAccount = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                                enPMSWorkorder.msdyn_BillingAccount = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                                jobId = service.Create(enPMSWorkorder);

                                entity["hil_jobid"] = new EntityReference("msdyn_workorder", jobId);
                                entity["hil_jobstatus"] = true;
                                entity["hil_description"] = "Job created successfully.";
                            }
                            else
                            {
                                entity["hil_description"] = "Product Sub Category Configuration (MG4) is missing.";
                                entity["hil_jobstatus"] = false;
                            }
                        }
                        else
                        {
                            if (entConsumer == null)
                            {
                                entity["hil_description"] = "Customer does not exist.";
                                entity["hil_jobstatus"] = false;
                            }
                            else if (erAddress == null)
                            {
                                entity["hil_description"] = addressRemarks + " Customer Address does not exist.";
                                entity["hil_jobstatus"] = false;
                            }
                        }
                    }
                    else
                    {
                        entity["hil_description"] = "Duplicate Job. Old Job# " + _duplicateWOIds;
                        entity["hil_jobstatus"] = false;
                    }
                }
                else
                {
                    entity["hil_description"] = strInputsValidationsummary;
                    entity["hil_jobstatus"] = false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.WorkOrderBulkUpload.PostCreate.Execute***  " + ex.Message);
            }
        }
        public static EntityReference GetCustomerAddress(IOrganizationService service, Guid customerGuid, Guid _pincode)
        {
            EntityReference retValue = null;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                LinkEntity _lnkEntity = new LinkEntity()
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "hil_address",
                    LinkFromAttributeName = "hil_address",
                    LinkToAttributeName = "hil_addressid",
                    Columns = new ColumnSet(false),
                };
                _lnkEntity.LinkCriteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, _pincode));

                Query = new QueryExpression("msdyn_workorder")
                {
                    ColumnSet = new ColumnSet("hil_address"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal, customerGuid));
                Query.TopCount = 1;
                Query.AddOrder("createdon", OrderType.Descending);
                if (_pincode != Guid.Empty)
                {
                    Query.LinkEntities.Add(_lnkEntity);
                }
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = enCol.Entities[0].GetAttributeValue<EntityReference>("hil_address");
                }
                else
                {
                    Query = new QueryExpression("hil_address")
                    {
                        ColumnSet = new ColumnSet("hil_fulladdress"),
                        Criteria = new FilterExpression(LogicalOperator.And)
                    };
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, customerGuid));
                    Query.TopCount = 1;
                    Query.AddOrder("createdon", OrderType.Descending);
                    enCol = service.RetrieveMultiple(Query);
                    if (enCol.Entities.Count > 0)
                    {
                        retValue = enCol.Entities[0].ToEntityReference();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.WorkOrderBulkUpload.PreCreate.GetCustomerAddress***  " + ex.Message);
            }
            return retValue;
        }
        public static EntityReference GetNatureOfComplaint(IOrganizationService service, string noc, Guid relatedProduct)
        {
            EntityReference retValue = null;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                Query = new QueryExpression("hil_natureofcomplaint")
                {
                    ColumnSet = new ColumnSet("hil_natureofcomplaintid"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, relatedProduct));
                Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, noc));
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Query.TopCount = 1;
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = enCol.Entities[0].ToEntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.WorkOrderBulkUpload.PreCreate.GetNatureOfComplaint***  " + ex.Message);
            }
            return retValue;
        }
        public static Entity ExecuteScalar(IOrganizationService service, string entityName, string primaryField, string primaryFieldValue, string[] columns, ref Entity entity)
        {
            Entity retEntity = null;
            try
            {
                QueryExpression Query = new QueryExpression(entityName);
                Query.ColumnSet = new ColumnSet(columns);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression(primaryField, ConditionOperator.Equal, primaryFieldValue));
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                EntityCollection enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count >= 1)
                {
                    retEntity = enCol.Entities[0];
                }
            }
            catch (Exception ex)
            {
                entity["hil_description"] = ex.Message;
                entity["hil_pmsjobstatus"] = false;
            }
            return retEntity;
        }
        public static bool GetCallSubTypeByNOC(IOrganizationService service, Guid callSubTypeGuid, Guid nocGuid)
        {
            bool retValue = false;
            QueryExpression Query;
            EntityCollection enCol;
            try
            {
                Query = new QueryExpression("hil_natureofcomplaint")
                {
                    ColumnSet = new ColumnSet("hil_natureofcomplaintid"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                Query.Criteria.AddCondition(new ConditionExpression("hil_natureofcomplaintid", ConditionOperator.Equal, nocGuid));
                Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, callSubTypeGuid));
                Query.TopCount = 1;
                enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    retValue = true;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.WorkOrderBulkUpload.PreCreate.GetCallSubTypeByNOC***  " + ex.Message);
            }
            return retValue;
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
        public static AMCBilling ValidateAMCReceiptAmount(AMCBilling _reqData, IOrganizationService service)
        {
            AMCBilling _retObj = new AMCBilling();
            string _fetchXML = string.Empty;
            DateTime _invoiceDate;

            try
            {
                if (_reqData.JobId == Guid.Empty)
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "Job Id is required." };
                    return _retObj;
                }
                if (_reqData.ReceiptAmount == null || _reqData.ReceiptAmount == 0)
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "Receipt Amount is required." };
                    return _retObj;
                }
                if (service != null)
                {
                    //if (_reqData.SourceCode == null)
                    //{
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorderincident'>
                        <attribute name='msdyn_name' />
                        <filter type='and'>
                          <condition attribute='msdyn_workorder' operator='eq' value='" + _reqData.JobId + @"' />
                            <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='ca'>
                            <attribute name='hil_invoicedate' />
                            </link-entity>
                            <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='wo'>
                            <attribute name='createdon' />
                            <attribute name='hil_actualcharges' />
                            <attribute name='hil_callsubtype' />
                            </link-entity>
                            </entity>
                            </fetch>";

                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        if (!entCol.Entities[0].Attributes.Contains("wo.hil_callsubtype"))
                        {
                            _retObj = new AMCBilling() { ResultStatus = false, ResultMessage = "Call Subtype is required." };
                            return _retObj;
                        }
                        if (!entCol.Entities[0].Attributes.Contains("ca.hil_invoicedate"))
                        {
                            _invoiceDate = new DateTime(1900, 1, 1);
                        }
                        else
                        {
                            _invoiceDate = (DateTime)(entCol.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_invoicedate").Value);
                        }
                        EntityReference entTemp = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_callsubtype").Value;

                        if (entTemp.Id != new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))
                        {
                            _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK." };
                            return _retObj;
                        }

                        decimal _payableAmount = 0;
                        if (entCol.Entities[0].Attributes.Contains("wo.hil_actualcharges"))
                        {
                            _payableAmount = ((Money)(entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_actualcharges").Value)).Value;
                        }
                        else
                        {
                            _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                        }
                        DateTime _jobDate = (DateTime)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.createdon").Value;
                        //_asOn Definition :: AMC Job Create date is concidered for Applying Discount rate becoz Product ageing also calculated from AMC Job Create Date
                        string _asOn = _jobDate.Year.ToString() + "-" + _jobDate.Month.ToString().PadLeft(2, '0') + "-" + _jobDate.Day.ToString().PadLeft(2, '0');
                        int _dayDiff = Convert.ToInt32(Math.Round((_jobDate - _invoiceDate).TotalDays, 0));
                        _dayDiff = _dayDiff < 0 ? 0 : _dayDiff;
                        decimal _stdDiscPer = 0;
                        decimal _spcDiscPer = 0;
                        decimal _stdDiscAmount = 0;
                        decimal _spcDiscAmount = 0;
                        //03B5A2D6-CC64-ED11-9562-6045BDAC526A - AMC Sale - FSM (Source)
                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_amcdiscountmatrix'>
                                <attribute name='hil_amcdiscountmatrixid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_discounttype' />
                                <attribute name='hil_discper' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_appliedto' operator='eq' value='{03B5A2D6-CC64-ED11-9562-6045BDAC526A}' />
                                    <condition attribute='hil_productaegingstart' operator='le' value='" + _dayDiff.ToString() + @"' />
                                    <condition attribute='hil_productageingend' operator='ge' value='" + _dayDiff.ToString() + @"' />
                                    <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                                    <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
                                </filter>
                                </entity>
                                </fetch>";

                        EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol1.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol1.Entities)
                            {
                                if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                                {
                                    _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                                }
                                else
                                {
                                    _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                                }
                            }
                            _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2); //Max Limit (90)
                            _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2); //Min Limit (85)
                            if (_reqData.ReceiptAmount >= _spcDiscAmount && _reqData.ReceiptAmount < _stdDiscAmount)
                            {
                                decimal _additionaldisc = Math.Round(_stdDiscAmount - Convert.ToDecimal(_reqData.ReceiptAmount), 2);
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "CONFIRMATION", ResultMessage = "To offer additional discount (Rs. " + _additionaldisc.ToString() + ") above Standard Discount, you need to take BSH approval. Click 'Yes' if approval already taken Or Click 'No'." };
                            }
                            else if (_reqData.ReceiptAmount < _spcDiscAmount)
                            {
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "As per AMC discount policy, you are allowed to collect minimum Rs. " + _stdDiscAmount.ToString() + "." };
                            }
                            else
                            {
                                _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                            }
                        }
                        else
                        {
                            if (_payableAmount != _reqData.ReceiptAmount)
                            {
                                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount." };
                            }
                            else
                            {
                                _retObj = new AMCBilling() { ResultStatus = true, ResultMessageType = "INFO", ResultMessage = "OK" };
                            }
                        }
                    }
                    else
                    {
                        _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "No Work Order Incident found." };
                    }
                    return _retObj;
                }
                else
                {
                    _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = "D365 Service Unavailable" };
                    return _retObj;
                }
            }
            catch (Exception ex)
            {
                _retObj = new AMCBilling() { ResultStatus = false, ResultMessageType = "WARNING", ResultMessage = ex.Message };
                return _retObj;
            }
        }
        public static IoTServiceCallRegistration IoTCreateServiceCallDealerPortal(IOrganizationService service, IoTServiceCallRegistration serviceCalldata)
        {
            IoTServiceCallRegistration objServiceCall;
            Guid customerGuid = Guid.Empty;
            Guid callSubTypeGuid = Guid.Empty;
            Guid serviceCallGuid = Guid.Empty;
            Entity lookupObj = null;
            EntityCollection entcoll;
            QueryExpression Query;
            string customerFullName = string.Empty;
            string customerMobileNumber = string.Empty;
            string customerEmail = string.Empty;
            Guid customerAssetGuid = Guid.Empty;
            DateTime? invoiceDate = null;
            string modelName = string.Empty;
            EntityReference erProductsubcategorymapping = null;
            EntityReference erProductCategory = null;
            EntityReference erProductsubcategory = null;

            EntityReference erNatureOfComplaint = null;
            EntityReference erCustomerAsset = null;
            bool continueFlag = false;
            string fullAddress = string.Empty;
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (serviceCalldata.CustomerMobleNo == string.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Nobile Number is required." };
                    }
                    if (serviceCalldata.CustomerGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Guid is required." };
                    }
                    if (serviceCalldata.NOCGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
                    }
                    if (serviceCalldata.AddressGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address is required." };
                    }
                    if (serviceCalldata.ProductCategoryGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required." };
                    }
                    if (serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Sub Category is required." };
                    }

                    Query = new QueryExpression("hil_address");
                    Query.ColumnSet = new ColumnSet("hil_addressid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, serviceCalldata.AddressGuid);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Address does not belong to Customer." };
                    }

                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, serviceCalldata.CustomerMobleNo);
                    Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer/Mobile No. does not exist." };
                    }
                    else
                    {
                        lookupObj = entcoll.Entities[0];
                        customerFullName = lookupObj.GetAttributeValue<string>("fullname");
                        customerEmail = lookupObj.GetAttributeValue<string>("emailaddress1");
                        customerMobileNumber = lookupObj.GetAttributeValue<string>("mobilephone"); // N
                    }
                    if (serviceCalldata.ChiefComplaint == string.Empty || serviceCalldata.ChiefComplaint == null || serviceCalldata.ChiefComplaint.Trim().Length == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer's Chief Complaint is required." };
                    }
                    //Case 1 Serial Number Exists
                    if (serviceCalldata.AssetGuid != Guid.Empty)
                    {
                        Entity ent = service.Retrieve("msdyn_customerasset", serviceCalldata.AssetGuid, new ColumnSet("msdyn_name", "hil_customer", "hil_productsubcategorymapping", "hil_productcategory", "hil_productsubcategory", "msdyn_customerassetid", "hil_invoicedate"));
                        if (ent != null)
                        {
                            erCustomerAsset = ent.ToEntityReference();
                            modelName = ent.GetAttributeValue<string>("msdyn_name");
                            invoiceDate = ent.GetAttributeValue<DateTime>("hil_invoicedate");
                            erProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory");
                            erProductsubcategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory");
                            erProductsubcategorymapping = ent.GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
                            continueFlag = true;
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Asset Serial Number does not exist." };
                        }
                    }
                    //Case 2 Product Category 
                    else if (serviceCalldata.ProductCategoryGuid != Guid.Empty)
                    {
                        erProductCategory = new EntityReference("product", serviceCalldata.ProductCategoryGuid);
                        erProductsubcategory = new EntityReference("product", serviceCalldata.ProductSubCategoryGuid);
                        modelName = string.Empty;
                        continueFlag = true;
                    }

                    if (!continueFlag)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Asset Serial Number/Product Category is required to proceed." };
                    }

                    if (serviceCalldata.SourceOfJob != 6)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Source of Job!!!" };
                    }

                    #region Get Nature of Complaint
                    string fetchXML = string.Empty;

                    if (serviceCalldata.NOCGuid != Guid.Empty)
                    {
                        fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                        fetchXML += "<entity name='hil_natureofcomplaint'>";
                        fetchXML += "<attribute name='hil_callsubtype' />";
                        fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                        fetchXML += "<order attribute='createdon' descending='false' />";
                        fetchXML += "<filter type='and'>";
                        fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                        fetchXML += "<condition attribute='hil_natureofcomplaintid' operator='eq' value='{" + serviceCalldata.NOCGuid + "}' />";
                        fetchXML += "</filter>";
                        fetchXML += "</entity>";
                        fetchXML += "</fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            erNatureOfComplaint = entcoll.Entities[0].ToEntityReference();
                            if (entcoll.Entities[0].Attributes.Contains("hil_callsubtype"))
                            {
                                callSubTypeGuid = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                            }
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "NOC does not match with Product Sub Category." };
                        }
                    }
                    #endregion
                    #region Create Service Call
                    objServiceCall = new IoTServiceCallRegistration();
                    objServiceCall = serviceCalldata;
                    Entity enWorkorder = new Entity("msdyn_workorder");

                    if (serviceCalldata.CustomerGuid != Guid.Empty)
                    {
                        enWorkorder["hil_customerref"] = new EntityReference("contact", serviceCalldata.CustomerGuid);
                    }
                    enWorkorder["hil_customername"] = customerFullName;
                    enWorkorder["hil_mobilenumber"] = customerMobileNumber;
                    enWorkorder["hil_email"] = customerEmail;

                    if (serviceCalldata.PreferredPartOfDay > 0 && serviceCalldata.PreferredPartOfDay < 4)
                    {
                        enWorkorder["hil_preferredtime"] = new OptionSetValue(serviceCalldata.PreferredPartOfDay);
                    }

                    if (serviceCalldata.PreferredDate != null && serviceCalldata.PreferredDate.Trim().Length > 0)
                    {
                        string _date = serviceCalldata.PreferredDate;
                        DateTime dtInvoice = new DateTime(Convert.ToInt32(_date.Substring(6, 4)), Convert.ToInt32(_date.Substring(0, 2)), Convert.ToInt32(_date.Substring(3, 2)));
                        enWorkorder["hil_preferreddate"] = dtInvoice;
                    }

                    if (serviceCalldata.AddressGuid != Guid.Empty)
                    {
                        enWorkorder["hil_address"] = new EntityReference("hil_address", serviceCalldata.AddressGuid);
                    }

                    if (erCustomerAsset != null)
                    {
                        enWorkorder["msdyn_customerasset"] = erCustomerAsset;
                    }

                    if (modelName != string.Empty)
                    {
                        enWorkorder["hil_modelname"] = modelName;
                    }

                    if (erProductCategory != null)
                    {
                        enWorkorder["hil_productcategory"] = erProductCategory;
                    }
                    if (erProductsubcategory != null)
                    {
                        enWorkorder["hil_productsubcategory"] = erProductsubcategory;
                    }

                    Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                    Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, erProductCategory.Id);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, erProductsubcategory.Id);
                    EntityCollection ec = service.RetrieveMultiple(Query);
                    if (ec.Entities.Count > 0)
                    {
                        enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
                    }

                    EntityCollection entCol;
                    Query = new QueryExpression("hil_consumertype");
                    Query.ColumnSet = new ColumnSet("hil_consumertypeid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
                    entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_consumertype"] = entCol.Entities[0].ToEntityReference();
                    }

                    Query = new QueryExpression("hil_consumercategory");
                    Query.ColumnSet = new ColumnSet("hil_consumercategoryid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
                    entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_consumercategory"] = entCol.Entities[0].ToEntityReference();
                    }

                    if (erNatureOfComplaint != null)
                    {
                        enWorkorder["hil_natureofcomplaint"] = erNatureOfComplaint;
                    }
                    if (callSubTypeGuid != Guid.Empty)
                    {
                        enWorkorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", callSubTypeGuid);
                    }
                    enWorkorder["hil_quantity"] = 1;
                    enWorkorder["hil_customercomplaintdescription"] = serviceCalldata.ChiefComplaint;
                    enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

                    enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); // SourceofJob:[{"6": "Dealer Portal"}]

                    enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                    enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                    string _dealerInfo = (serviceCalldata.DealerCode + "|" + serviceCalldata.DealerName);
                    enWorkorder["hil_newserialnumber"] = _dealerInfo.Length > 99 ? _dealerInfo.Substring(0, 99) : _dealerInfo; //Dealer Code/Name 

                    serviceCallGuid = service.Create(enWorkorder);
                    if (serviceCallGuid != Guid.Empty)
                    {
                        if (serviceCalldata.DealerCode != null && serviceCalldata.DealerCode != "" && serviceCalldata.DealerCode.Trim().Length > 0)
                        {
                            Query = new QueryExpression("hil_jobsextension");
                            Query.ColumnSet = new ColumnSet("hil_dealercode", "hil_dealername");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_jobs", ConditionOperator.Equal, serviceCallGuid);
                            EntityCollection entColExt = service.RetrieveMultiple(Query);
                            if (entColExt.Entities.Count > 0)
                            {
                                Entity _entExt = entColExt.Entities[0];
                                _entExt["hil_dealercode"] = serviceCalldata.DealerCode;
                                _entExt["hil_dealername"] = serviceCalldata.DealerName;
                                try
                                {
                                    service.Update(_entExt);
                                }
                                catch (Exception ex)
                                {
                                    objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                                    return objServiceCall;
                                }
                            }
                        }
                        objServiceCall.JobGuid = serviceCallGuid;
                        objServiceCall.JobId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
                        objServiceCall.StatusCode = "200";
                        objServiceCall.StatusDescription = "OK";
                    }
                    else
                    {
                        objServiceCall.StatusCode = "204";
                        objServiceCall.StatusDescription = "FAILURE !!! Something went wrong";
                    }
                    return objServiceCall;
                    #endregion
                }
                else
                {
                    objServiceCall = new IoTServiceCallRegistration { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    return objServiceCall;
                }
            }
            catch (Exception ex)
            {
                objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                return objServiceCall;
            }
        }
    }
    public class RequestDataCallMasking1
    {
        public string MobileNumber { get; set; }

        public string JobNumber { get; set; }
    }
    public class IoTServiceCallRegistration
    {
        public string SerialNumber { get; set; }
        public string ProductModelNumber { get; set; }
        public Guid NOCGuid { get; set; }
        public string NOCName { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public Guid ProductSubCategoryGuid { get; set; }
        public string ChiefComplaint { get; set; }
        public Guid AddressGuid { get; set; }
        public Guid AssetGuid { get; set; }
        public string CustomerMobleNo { get; set; }
        public Guid CustomerGuid { get; set; }
        public Guid JobGuid { get; set; }
        public string JobId { get; set; }
        public string ImageBase64String { get; set; }
        public int ImageType { get; set; }
        public int SourceOfJob { get; set; }
        public string PreferredDate { get; set; }
        public int PreferredPartOfDay { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string DealerCode { get; set; }
        public string DealerName { get; set; }
        public string CustomerName { get; set; }
        public string AddressLine1 { get; set; }
        public string Pincode { get; set; }
        public string PreferredLanguage { get; set; }
    }

    public class AMCBilling
    {
        public Guid JobId { get; set; }

        public Decimal? ReceiptAmount { get; set; }

        public int? SourceCode { get; set; }

        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
        public string ResultMessageType { get; set; }
    }
}
