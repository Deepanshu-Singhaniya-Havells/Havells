using System;
using System.Globalization;
using HavellsNewPlugin.TenderModule.MailtoTeams;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.DOA_Approval
{
    public class SendMailToApprover : IPlugin
    {
        public static ITracingService tracingService = null;
        public string _primaryField = string.Empty;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                   && context.PrimaryEntityName.ToLower() == "hil_oaheader" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    tracingService.Trace("Send Approval Email plugin started");
                    //var entityName = "hil_oaheader";// context.InputParameters["EntityName"].ToString();
                    //var entityId = "9745ff20-23a0-ec11-b400-6045bdaac501";// context.InputParameters["EntityID"].ToString();
                    //tracingService.Trace("entityName      " + entityName);

                    QueryExpression query = new QueryExpression(entity.LogicalName);
                    query.ColumnSet = new ColumnSet("hil_customercode", "hil_customername", "hil_creditdays", "hil_salesoffice", "hil_name",
                        "hil_expecteddateofcollection", "hil_ordervalue", "hil_tenderid", "hil_outstandingamount", "hil_creditlimit", "hil_tolerance",
                        "hil_requestremarks", "hil_limitrequested", "hil_buhead", "hil_approvalstatus", "hil_department");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression(entity.LogicalName + "id", ConditionOperator.Equal, entity.Id));
                    LinkEntity EntityA = new LinkEntity("hil_oaheader", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                    EntityA.Columns = new ColumnSet("hil_projectname", "hil_salesoffice", "hil_paymentterm", "hil_bankguaranteerequired", "hil_categoryofbankguarantee"
                        , "hil_valueofbg", "hil_warrantyperiod", "hil_zonalhead");
                    EntityA.EntityAlias = "OCL";
                    query.LinkEntities.Add(EntityA);

                    EntityCollection entColl = service.RetrieveMultiple(query);
                    if (entColl.Entities.Count == 1)
                    {
                        Entity _OAHeader = entColl.Entities[0];
                        if (_OAHeader.Contains("hil_approvalstatus") && _OAHeader.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value == 3)
                        {

                            //    Entity oaUpdate = new Entity(_OAHeader.LogicalName);
                            //    oaUpdate.Id = _OAHeader.Id;
                            //    oaUpdate["hil_approvalstatus"] = new OptionSetValue(3);
                            //    service.Update(oaUpdate);

                            // throw new InvalidPluginExecutionException("ADDPPDPDPD");

                            Entity app = service.Retrieve("hil_approvalmatrix", new Guid("45088ec7-e1a4-ec11-9840-6045bdac66ea"), new ColumnSet("hil_mailbody"));

                            string mailbody = genrateMailBody(_OAHeader, app.GetAttributeValue<string>("hil_mailbody"));

                            string projectName = _OAHeader.Contains("OCL.hil_projectname") ? _OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_projectname").Value.ToString() : "";//
                            string customername = _OAHeader.Contains("hil_customername") ? _OAHeader.GetAttributeValue<EntityReference>("hil_customername").Name : "";
                            string salesOfficeName = _OAHeader.Contains("hil_salesoffice") ? _OAHeader.GetAttributeValue<EntityReference>("hil_salesoffice").Name : "";//
                            string orderValue = _OAHeader.Contains("hil_ordervalue") ? String.Format("{0:0.00}", String.Format(new CultureInfo("en-IN", false), "{0:n}",
                                _OAHeader.GetAttributeValue<Money>("hil_ordervalue").Value)) : "";//

                            string subject = @"APPROVAL(Cable BU) FOR LIMIT ON SELF EXPOSURE, " + customername + @", " + salesOfficeName + @" ₹ " + orderValue;

                            Entity entEmail = new Entity("email");

                            Entity entFrom = new Entity("activityparty");
                            entFrom["partyid"] = getSender("SMS", service);
                            Entity[] entFromList = { entFrom };
                            entEmail["from"] = entFromList;

                            EntityReference to = new EntityReference("hil_bdteammember", new Guid("cbef6941-9c94-ec11-b400-0022486edff5"));//CMD
                            Entity toActivityParty = new Entity("activityparty");
                            toActivityParty["partyid"] = to;
                            entEmail["to"] = new Entity[] { toActivityParty };


                            EntityCollection entCCList = new EntityCollection();
                            entCCList.EntityName = "systemuser";

                            EntityReference department = _OAHeader.GetAttributeValue<EntityReference>("hil_department");

                            EntityCollection toTeamMembers = new EntityCollection();
                            QueryExpression _query = new QueryExpression("hil_bdteam");
                            _query.ColumnSet = new ColumnSet(false);
                            _query.Criteria = new FilterExpression(LogicalOperator.And);
                            _query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "DOA CC"));
                            _query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, department.Id));
                            _query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                            EntityCollection bdteamCol = service.RetrieveMultiple(_query);
                            tracingService.Trace("bdteamCol.Entities.Count " + bdteamCol.Entities.Count);
                            Entity entCC = new Entity("activityparty");
                            if (bdteamCol.Entities.Count > 0)
                            {
                                tracingService.Trace("bdteamCol count " + bdteamCol.Entities.Count);

                                QueryExpression _querymem = new QueryExpression("hil_bdteammember");
                                _querymem.ColumnSet = new ColumnSet("emailaddress");
                                _querymem.Criteria = new FilterExpression(LogicalOperator.And);
                                _querymem.Criteria.AddCondition(new ConditionExpression("hil_team", ConditionOperator.Equal, bdteamCol.Entities[0].Id));
                                _querymem.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                                EntityCollection bdteammemCol = service.RetrieveMultiple(_querymem);
                                EntityCollection entTOList = new EntityCollection();

                                if (bdteammemCol.Entities.Count > 0)
                                {
                                    foreach (Entity entity1 in bdteammemCol.Entities)
                                    {
                                        tracingService.Trace("toTeamMembers " + entity1.ToEntityReference());
                                        toTeamMembers.Entities.Add(entity1);
                                        string name = (string)entity1.GetAttributeValue<string>("emailaddress");
                                        tracingService.Trace("emailaddress" + name);
                                        entCC = new Entity("activityparty");
                                        entCC["partyid"] = entity1.ToEntityReference();
                                        entCCList.Entities.Add(entCC);
                                    }
                                }
                            }

                            //  entCCList = SendEmailtoteamsonOA.retriveTeamMembers(service, "DOA CC", null, department, null, entCCList);

                          

                            //entCC = new Entity("activityparty");
                            //entCC["partyid"] = new EntityReference("hil_bdteammember", new Guid("3dacf02c-9aa9-ec11-9840-6045bdad4f4b"));//Director Finance
                            //entCCList.Entities.Add(entCC);

                            //entCC = new Entity("activityparty");
                            //entCC["partyid"] = new EntityReference("systemuser", new Guid("4d0c17d9-051b-e911-a954-000d3af06a16"));//Narender singh CMT
                            //entCCList.Entities.Add(entCC);

                            //entCC = new Entity("activityparty");
                            //entCC["partyid"] = new EntityReference("systemuser", new Guid("82830c91-a2d1-ea11-a813-000d3af0501c"));//Harpreet3.Singh
                            //entCCList.Entities.Add(entCC);

                            if (_OAHeader.Contains("OCL.hil_zonalhead"))
                            {
                                entCC = new Entity("activityparty");
                                entCC["partyid"] = (EntityReference)_OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_zonalhead").Value;
                                entCCList.Entities.Add(entCC);
                            }

                            entCC = new Entity("activityparty");
                            entCC["partyid"] = _OAHeader.GetAttributeValue<EntityReference>("hil_buhead");
                            entCCList.Entities.Add(entCC);

                            //Entity ccActivityParty = new Entity("activityparty");
                            //ccActivityParty["partyid"] = 

                            entEmail["cc"] = entCCList;

                            entEmail["subject"] = subject;
                            entEmail["description"] = mailbody;

                            entEmail["regardingobjectid"] = _OAHeader.ToEntityReference();

                            Guid emailId = service.Create(entEmail);

                            SendEmailRequest sendEmailReq = new SendEmailRequest()
                            {
                                EmailId = emailId,
                                IssueSend = true
                            };
                            SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
                        }
                        else
                        {
                            //throw new InvalidPluginExecutionException("AssssssssssssssDDPPDPDPD");
                        }
                    }
                    else
                    {
                        // throw new InvalidPluginExecutionException("AssssssssssssssDDPPDqqqqqqqqqqqqPDPD");
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Send Email Error:- " + ex.Message);
            }
        }
        public static string genrateMailBody(Entity _OAHeader, string mailbody)
        {
            string projectName = _OAHeader.Contains("OCL.hil_projectname") ? _OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_projectname").Value.ToString() : "";//
            string salesOfficeName = _OAHeader.Contains("hil_salesoffice") ? _OAHeader.GetAttributeValue<EntityReference>("hil_salesoffice").Name : "";//
            string paymentTerms = _OAHeader.Contains("OCL.hil_paymentterm") ? _OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_paymentterm").Value.ToString() : "";//
                                                                                                                                                                     // string pbgCondition = _OAHeader.Contains("OCL.hil_bankguaranteerequired") ? ((bool)_OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_bankguaranteerequired").Value ? "Applicable" : "Not Applicable") : "";

            int categoryofBgValue = _OAHeader.Contains("OCL.hil_categoryofbankguarantee") ? ((OptionSetValue)_OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_categoryofbankguarantee").Value).Value : 0;
            string categoryofBg = categoryofBgValue == 1 ? "Security" : categoryofBgValue == 2 ? "Advance" : categoryofBgValue == 3 ? "Performance" : categoryofBgValue == 4 ? "Retension" : "";

            int valueofBgValue = _OAHeader.Contains("OCL.hil_valueofbg") ? ((OptionSetValue)_OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_valueofbg").Value).Value : 0;
            string valueofBg = valueofBgValue == 910590000 ? "0%" : valueofBgValue == 910590001 ? "5%" : valueofBgValue == 910590002 ? "10%" :
                valueofBgValue == 910590003 ? "15%" : valueofBgValue == 910590004 ? "20%" : "";

            string warrentyPeriod = _OAHeader.Contains("OCL.hil_warrantyperiod") ?
                ((EntityReference)_OAHeader.GetAttributeValue<AliasedValue>("OCL.hil_warrantyperiod").Value).Name :
                "";

            string pbgCondition = categoryofBg + ", " + valueofBg + ", " + warrentyPeriod;

            string oaNumber = _OAHeader.GetAttributeValue<string>("hil_name");

            string orderValue = _OAHeader.Contains("hil_ordervalue") ?
               String.Format("{0:0.00}", String.Format(new CultureInfo("en-IN", false), "{0:n}", _OAHeader.GetAttributeValue<Money>("hil_ordervalue").Value)) : "";//
            string customerCode = _OAHeader.Contains("hil_customercode") ? _OAHeader.GetAttributeValue<string>("hil_customercode") : "";//
            string tenderNo = _OAHeader.Contains("hil_tenderid") ? _OAHeader.GetAttributeValue<EntityReference>("hil_tenderid").Name : "";//

            string customername = _OAHeader.Contains("hil_customername") ? _OAHeader.GetAttributeValue<EntityReference>("hil_customername").Name : "";
            string outstandingAmount = _OAHeader.Contains("hil_outstandingamount") ?
                String.Format("{0:0.00}", String.Format(new CultureInfo("en-IN", false), "{0:n}", _OAHeader.GetAttributeValue<Money>("hil_outstandingamount").Value)) : "";
            string overdueDays = _OAHeader.Contains("hil_creditdays") ? _OAHeader.GetAttributeValue<int>("hil_creditdays").ToString() : "";

            string expectedDate = _OAHeader.Contains("hil_expecteddateofcollection") ? DateTimeToString(_OAHeader.GetAttributeValue<DateTime>("hil_expecteddateofcollection")) : "";

            string tolerance = _OAHeader.Contains("hil_tolerance") ? String.Format("{0:0.00}", _OAHeader.GetAttributeValue<decimal>("hil_tolerance")) : "";
            string BUremarks = _OAHeader.Contains("hil_requestremarks") ? _OAHeader.GetAttributeValue<string>("hil_requestremarks") : "";
            string limitRequested = _OAHeader.Contains("hil_limitrequested") ?
               String.Format("{0:0.00}", String.Format(new CultureInfo("en-IN", false), "{0:n}", _OAHeader.GetAttributeValue<Money>("hil_limitrequested").Value)) : "";

            mailbody = mailbody.Replace("{{Subject}}", "LIMIT REQUEST FOR CUSTOMER (" + customerCode + ") IS PENDING FOR APPROVAL");

            //[10:41] NagmaniNath Tiwari
            //            Dear Sir,
            //          Cable team(< Branch Name > sales office) has received an order from Hetero Drugs Limited for Bonthapally - Unit - IV.Project.

            mailbody = mailbody.Replace("{{BodyHeader}}", "Dear Sir, <br>Cable team (" + salesOfficeName + ") has received an order from  " + customername + "  for " + projectName + " Project");

            mailbody = mailbody.Replace("{{CustomerCode}}", customerCode);

            mailbody = mailbody.Replace("{{TenderNo}}", tenderNo);

            mailbody = mailbody.Replace("{{OANumber}}", oaNumber);

            mailbody = mailbody.Replace("{{PBGCondition}}", pbgCondition);

            mailbody = mailbody.Replace("{{EDC}}", expectedDate);

            mailbody = mailbody.Replace("{{Tolerance}}", tolerance);

            mailbody = mailbody.Replace("{{LimitRequested}}", limitRequested);

            mailbody = mailbody.Replace("{{OutstandingAmount}}", outstandingAmount);

            mailbody = mailbody.Replace("{{OverdueDays}}", overdueDays);

            mailbody = mailbody.Replace("{{PaymentTerms}}", paymentTerms);

            mailbody = mailbody.Replace("{{BURemarks}}", BUremarks);

            mailbody = mailbody.Replace("{{RejectedSubject}}", "OA " + oaNumber + " | LIMIT REQUEST FOR CUSTOMER | Rejected");

            mailbody = mailbody.Replace("{{ApprovedSubject}}", "OA " + oaNumber + " | LIMIT REQUEST FOR CUSTOMER | Approved");
            return (mailbody);
        }
        public static string DateTimeToString(DateTime _cTimeStamp)
        {
            tracingService.Trace("DateTimeToString Function Started ");
            string timestamp = null;
            timestamp = _cTimeStamp.Day.ToString().PadLeft(2, '0') + "-" + _cTimeStamp.Month.ToString().PadLeft(2, '0') + "-" + _cTimeStamp.Year.ToString(); ;
            tracingService.Trace("timestamp  " + timestamp);
            tracingService.Trace("DateTimeToString Function Ended");
            return timestamp;
        }
        public static EntityReference getSender(string queName, IOrganizationService service)
        {
            EntityReference sender = new EntityReference();
            QueryExpression _queQuery = new QueryExpression("queue");
            _queQuery.ColumnSet = new ColumnSet("name");
            _queQuery.Criteria = new FilterExpression(LogicalOperator.And);
            _queQuery.Criteria.AddCondition("name", ConditionOperator.Equal, queName);
            EntityCollection queueColl = service.RetrieveMultiple(_queQuery);
            if (queueColl.Entities.Count == 1)
            {
                sender = queueColl[0].ToEntityReference();
            }
            return sender;
        }
    }
}
