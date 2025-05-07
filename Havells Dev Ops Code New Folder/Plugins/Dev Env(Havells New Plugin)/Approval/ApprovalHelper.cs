using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace HavellsNewPlugin.Approval
{
    public class ApprovalHelper
    {
        const String URL = @"https://havells.crm8.dynamics.com/main.aspx?appid=675ffa44-b6d0-ea11-a813-000d3af05d7b&forceUCI=1&pagetype=entityrecord&etn=";
        public static void createApproval(string entityName, string entityId, string purpose, string _primaryField, IOrganizationService service, ITracingService tracingService)
        {
            try
            {

                tracingService.Trace("createApproval method of Approval Helper is Started");

                QueryExpression _query = new QueryExpression("hil_approvalmatrix");
                _query.ColumnSet = new ColumnSet("hil_approvalmatrixid", "hil_name", "createdon", "hil_kpi",
                    "statecode", "hil_purpose", "hil_mailbody", "hil_level", "hil_entity", "hil_approver", "hil_discountmin", "hil_discountmax",
                    "hil_approverposition", "hil_copytoposition", "hil_duehrs", "hil_optional");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_entity", ConditionOperator.Equal, entityName));
                _query.Criteria.AddCondition(new ConditionExpression("hil_purpose", ConditionOperator.Equal, purpose));
                _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                _query.AddOrder("hil_level", OrderType.Ascending);
                EntityCollection ApprovalMatrixColl = service.RetrieveMultiple(_query);

                Guid lastApproval = Guid.Empty;
                Guid currentApproval = Guid.Empty;
                if (entityName == "hil_orderchecklist")
                {

                    Entity ocl = service.Retrieve(entityName, new Guid(entityId), new ColumnSet(true));
                    Decimal maxDiscoount = getMaxDiscount(ocl, service, tracingService);
                    
                    tracingService.Trace("getMaxDiscount ended value is " + maxDiscoount);

                    Decimal maxApprovalMatrixDiscont = getMaxDiscountFromApprovalMatrix(service, purpose, tracingService);
                    if (maxDiscoount > maxApprovalMatrixDiscont)
                    {
                        throw new InvalidPluginExecutionException("createApproval Error: Discount is crossed the approved limit.");
                    }

                    for (int i = 0; i < ApprovalMatrixColl.Entities.Count; i++) // Entity entColl[i] in entColl.Entities)
                    {
                        decimal max = ApprovalMatrixColl[i].GetAttributeValue<decimal>("hil_discountmax");
                        tracingService.Trace("getMaxDiscount ended value is " + max);

                        decimal min = ApprovalMatrixColl[i].GetAttributeValue<decimal>("hil_discountmin");
                        tracingService.Trace("getMaxDiscount ended value is " + min);
                        
                        if (maxDiscoount >= min)
                        {
                            ColumnSet collSet = findEntityColl(ApprovalMatrixColl[i].GetAttributeValue<string>("hil_mailbody"), _primaryField, service, tracingService);
                            Entity target = service.Retrieve(entityName, new Guid(entityId), collSet);
                            EntityReference targetOwner = target.GetAttributeValue<EntityReference>("ownerid");

                            _query = new QueryExpression("hil_userbranchmapping");
                            _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead", "hil_scm");
                            _query.Criteria = new FilterExpression(LogicalOperator.And);
                            _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
                            _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

                            EntityCollection userMapingColl = service.RetrieveMultiple(_query);

                            Entity _approval = new Entity("hil_approval");
                            EntityReference approver = new EntityReference();
                            _approval["subject"] = ApprovalMatrixColl[i].GetAttributeValue<string>("hil_purpose") + "_Level " + ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level");

                            tracingService.Trace("levrl");
                            if (ApprovalMatrixColl[i].Contains("hil_approver"))
                                approver = ApprovalMatrixColl[i].GetAttributeValue<EntityReference>("hil_approver");
                            else if (ApprovalMatrixColl[i].Contains("hil_approverposition"))
                            {
                                String approvalPosition = ApprovalMatrixColl[i].GetAttributeValue<EntityReference>("hil_approverposition").Name;
                                if (userMapingColl.Entities.Count > 0)
                                    approver = getApproverByPosition(userMapingColl[0], approvalPosition, service, tracingService);
                                else
                                    throw new InvalidPluginExecutionException("Approver not Found");

                            }
                            _approval["ownerid"] = approver;
                            _approval["hil_level"] = new OptionSetValue(ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level"));
                            _approval["regardingobjectid"] = target.ToEntityReference();// new EntityReference(entityName, new Guid(entityId));
                            if ((ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level") == 1))
                            {
                                _approval["hil_approvalstatus"] = new OptionSetValue(3);// 3 for - submit for approval    4 for - Draft
                                _approval["hil_requesteddate"] = DateTime.Now.AddMinutes(330);
                            }
                            else
                            {
                                _approval["hil_approvalstatus"] = new OptionSetValue(4);// 4 for - Draft
                            }
                            //if (ApprovalMatrixColl[i].Contains("hil_optional"))
                            //{
                            //    _approval["hil_optional"] = ApprovalMatrixColl[i].GetAttributeValue<bool>("hil_optional");

                            //}
                            if (ApprovalMatrixColl[i].Contains("hil_duehrs"))
                            {
                                _approval["scheduledend"] = DateTime.Now.AddHours(ApprovalMatrixColl[i].GetAttributeValue<int>("hil_duehrs")).AddMinutes(330);
                            }
                            currentApproval = service.Create(_approval);
                            if ((ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level") == 1))
                            {
                                // string subject = "Approval Required for " + purpose + " ID " + target[_primaryField];

                                EntityCollection entToList = new EntityCollection();
                                entToList.EntityName = "systemuser";
                                //String approvalPosition = ApprovalMatrixColl[i].GetAttributeValue<string>("hil_approverposition");
                                if (ApprovalMatrixColl[i].Contains("hil_copytoposition"))
                                {

                                    string copyto = ApprovalMatrixColl[i].GetAttributeValue<string>("hil_copytoposition");

                                    if (userMapingColl.Entities.Count > 0 && copyto.Contains(","))
                                    {
                                        entToList = getCopyToData(copyto, userMapingColl, service);
                                    }

                                }
                                Entity entTo = new Entity("activityparty");
                                entTo["partyid"] = targetOwner;
                                entToList.Entities.Add(entTo);

                                string mailbody = createEmailBody(target, ApprovalMatrixColl[0].GetAttributeValue<string>("hil_mailbody"), approver.Name, collSet, service, tracingService);
                                string subject = createEmailSubject(target, _primaryField, ApprovalMatrixColl[0].GetAttributeValue<string>("hil_kpi"), service, tracingService);
                                sendEmal(approver, entToList, target.ToEntityReference(), mailbody, subject, service);
                                lastApproval = currentApproval;


                            }
                            else
                            {
                                Entity entity = new Entity("hil_approval");
                                entity.Id = lastApproval;
                                entity["hil_nextapproval"] = new EntityReference(_approval.LogicalName, currentApproval);
                                //entity["hil_nextisoptional"] = ApprovalMatrixColl[i].GetAttributeValue<bool>("hil_optional");
                                service.Update(entity);
                                lastApproval = currentApproval;
                            }
                        }
                    }
                    //throw new InvalidPluginExecutionException("hhhh");
                }
                else
                {
                    for (int i = 0; i < ApprovalMatrixColl.Entities.Count; i++) // Entity entColl[i] in entColl.Entities)
                    {
                        ColumnSet collSet = findEntityColl(ApprovalMatrixColl[i].GetAttributeValue<string>("hil_mailbody"), _primaryField, service, tracingService);
                        Entity target = service.Retrieve(entityName, new Guid(entityId), collSet);
                        EntityReference targetOwner = target.GetAttributeValue<EntityReference>("ownerid");

                        _query = new QueryExpression("hil_userbranchmapping");
                        _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead", "hil_scm");
                        _query.Criteria = new FilterExpression(LogicalOperator.And);
                        _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
                        _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

                        EntityCollection userMapingColl = service.RetrieveMultiple(_query);

                        Entity _approval = new Entity("hil_approval");
                        EntityReference approver = new EntityReference();
                        _approval["subject"] = ApprovalMatrixColl[i].GetAttributeValue<string>("hil_purpose") + "_Level " + ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level");

                        tracingService.Trace("levrl");
                        if (ApprovalMatrixColl[i].Contains("hil_approver"))
                            approver = ApprovalMatrixColl[i].GetAttributeValue<EntityReference>("hil_approver");
                        else if (ApprovalMatrixColl[i].Contains("hil_approverposition"))
                        {
                            String approvalPosition = ApprovalMatrixColl[i].GetAttributeValue<EntityReference>("hil_approverposition").Name;
                            if (userMapingColl.Entities.Count > 0)
                                approver = getApproverByPosition(userMapingColl[0], approvalPosition, service, tracingService);
                            else
                                throw new InvalidPluginExecutionException("Approver not Found");

                        }
                        _approval["ownerid"] = approver;
                        _approval["hil_level"] = new OptionSetValue(ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level"));
                        _approval["regardingobjectid"] = target.ToEntityReference();// new EntityReference(entityName, new Guid(entityId));
                        if ((ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level") == 1))
                        {
                            _approval["hil_approvalstatus"] = new OptionSetValue(3);// 3 for - submit for approval    4 for - Draft
                            _approval["hil_requesteddate"] = DateTime.Now.AddMinutes(330);
                        }
                        else
                        {
                            _approval["hil_approvalstatus"] = new OptionSetValue(4);// 4 for - Draft
                        }
                        //if (ApprovalMatrixColl[i].Contains("hil_optional"))
                        //{
                        //    _approval["hil_optional"] = ApprovalMatrixColl[i].GetAttributeValue<bool>("hil_optional");

                        //}
                        if (ApprovalMatrixColl[i].Contains("hil_duehrs"))
                        {
                            _approval["scheduledend"] = DateTime.Now.AddHours(ApprovalMatrixColl[i].GetAttributeValue<int>("hil_duehrs")).AddMinutes(330);
                        }
                        currentApproval = service.Create(_approval);
                        if ((ApprovalMatrixColl[i].GetAttributeValue<int>("hil_level") == 1))
                        {
                            // string subject = "Approval Required for " + purpose + " ID " + target[_primaryField];

                            EntityCollection entToList = new EntityCollection();
                            entToList.EntityName = "systemuser";
                            //String approvalPosition = ApprovalMatrixColl[i].GetAttributeValue<string>("hil_approverposition");
                            if (ApprovalMatrixColl[i].Contains("hil_copytoposition"))
                            {

                                string copyto = ApprovalMatrixColl[i].GetAttributeValue<string>("hil_copytoposition");

                                if (userMapingColl.Entities.Count > 0 && copyto.Contains(","))
                                {
                                    entToList = getCopyToData(copyto, userMapingColl, service);
                                }

                            }
                            Entity entTo = new Entity("activityparty");
                            entTo["partyid"] = targetOwner;
                            entToList.Entities.Add(entTo);

                            string mailbody = createEmailBody(target, ApprovalMatrixColl[0].GetAttributeValue<string>("hil_mailbody"), approver.Name, collSet, service, tracingService);
                            string subject = createEmailSubject(target, _primaryField, ApprovalMatrixColl[0].GetAttributeValue<string>("hil_kpi"), service, tracingService);
                           // if (target.LogicalName == "hil_tenderbankguarantee")
                                sendEmal(approver, entToList, target.ToEntityReference(), mailbody, subject, service);
                            lastApproval = currentApproval;


                        }
                        else
                        {
                            Entity entity = new Entity("hil_approval");
                            entity.Id = lastApproval;
                            entity["hil_nextapproval"] = new EntityReference(_approval.LogicalName, currentApproval);
                            //entity["hil_nextisoptional"] = ApprovalMatrixColl[i].GetAttributeValue<bool>("hil_optional");
                            service.Update(entity);
                            lastApproval = currentApproval;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("|| createApproval Error: " + ex.Message);
            }
        }
        public static decimal getMaxDiscountFromApprovalMatrix(IOrganizationService service,string purpose, ITracingService tracing)
        {
            decimal maxvalue = 0;
            string fetchMaxDiscount = $@"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                                            <entity name='hil_approvalmatrix'>
                                                <attribute name='hil_purpose' alias='approvalmatrix' groupby='true' />
                                                <attribute name='hil_discountmax' alias='amount' aggregate='max' />
                                                <filter type='and'>
                                                    <condition attribute='hil_purpose' operator='eq' value='{purpose}' />
                                                    <condition attribute='statecode' operator='eq' value='0' />
                                                </filter>
                                            </entity>
                                        </fetch>";
            EntityCollection eColl = service.RetrieveMultiple(new FetchExpression(fetchMaxDiscount));



            maxvalue = (decimal)((AliasedValue)eColl.Entities[0]["amount"]).Value;



            return maxvalue;
        }
        public static decimal getMaxDiscount(Entity orderCheckList, IOrganizationService service, ITracingService tracing)
        {
            tracing.Trace("getMaxDiscount Started");
            string fetchPODiscount = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
	                            <entity name='hil_orderchecklistproduct'>
		                            <attribute name='hil_orderchecklistid' alias='orderchecklist' groupby='true' />
		                            <attribute name='hil_podiscount' alias='amount' aggregate='max' />
		                            <filter type='and'>
			                            <condition attribute='hil_orderchecklistid' operator='eq' value='" + orderCheckList.Id + @"' />
			                            <condition attribute='statecode' operator='eq' value='0' />
		                            </filter>
	                            </entity>
                            </fetch>";
            string fetchOrderBookingDiscount = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
	                            <entity name='hil_orderchecklistproduct'>
		                            <attribute name='hil_orderchecklistid' alias='orderchecklist' groupby='true' />
		                            <attribute name='hil_discount' alias='amount' aggregate='max' />
		                            <filter type='and'>
			                            <condition attribute='hil_orderchecklistid' operator='eq' value='" + orderCheckList.Id + @"' />
			                            <condition attribute='statecode' operator='eq' value='0' />
		                            </filter>
	                            </entity>
                            </fetch>";
            if (orderCheckList.Contains("hil_tenderno"))
            {
                EntityCollection eColl = service.RetrieveMultiple(new FetchExpression(fetchPODiscount));
                return (decimal)((AliasedValue)eColl.Entities[0]["amount"]).Value;
            }
            else
            {
                EntityCollection eColl = service.RetrieveMultiple(new FetchExpression(fetchOrderBookingDiscount));
                return (decimal)((AliasedValue)eColl.Entities[0]["amount"]).Value;
            }
        }
        public static EntityReference getApproverByPosition(Entity userMaping, string approvalPosition, IOrganizationService service, ITracingService tracingService)
        {
            tracingService.Trace("getApproverByPosition Stratrd");
            EntityReference approvr = new EntityReference();
            if (approvalPosition == "Branch Product Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_branchproducthead");
            }
            else if (approvalPosition == "Enquiry Creator")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_user");
            }
            else if (approvalPosition == "Zonal Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_zonalhead");
            }
            else if (approvalPosition == "BU Head")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_buhead");
            }
            else if (approvalPosition == "Sr. Manager Finance")
            {
                QueryExpression query = new QueryExpression("systemuser");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
                EntityCollection _entitys = service.RetrieveMultiple(query);

                if (_entitys.Entities.Count > 0)
                    approvr = _entitys[0].ToEntityReference();
            }
            else if (approvalPosition == "Sr. Manager Treasury")
            {
                QueryExpression query = new QueryExpression("user");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("{45C29172-3CD0-EB11-BACC-6045BD72E9C2}"));
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count > 0)
                    approvr = _entitys[0].ToEntityReference();
            }
            else if (approvalPosition == "SCM")
            {
                approvr = userMaping.GetAttributeValue<EntityReference>("hil_scm");
            }
            //else if (approvalPosition == "Design Head")
            //{
            //    QueryExpression query = new QueryExpression("systemuser");
            //    query.ColumnSet = new ColumnSet(true);
            //    query.Criteria = new FilterExpression(LogicalOperator.And);
            //    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
            //    EntityCollection _entitys = service.RetrieveMultiple(query);

            //    _approval["ownerid"] = entColl[i].GetAttributeValue<EntityReference>("hil_user"); = _entitys[0].ToEntityReference();
            //}
            tracingService.Trace("getApproverByPosition end");
            return approvr;
        }
        public static EntityCollection getCopyToData(string copyto, EntityCollection userMapingColl, IOrganizationService service)
        {
            EntityCollection entToList = new EntityCollection();
            entToList.EntityName = "systemuser";
            String[] mailto = copyto.Split(',');
            Entity entTo = new Entity("activityparty");
            foreach (string to in mailto)
            {
                if (to == "Branch Product Head")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_branchproducthead");

                }
                //else if (to == "Design Team")
                //{
                //    entTo = new Entity("activityparty");
                //    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_branchproducthead");
                //}
                else if (to == "Enquiry Creator")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_user"); ;
                }
                else if (to == "Zonal Head")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_zonalhead");
                }
                else if (to == "BU Head")
                {
                    entTo = new Entity("activityparty");
                    entTo["partyid"] = userMapingColl[0].GetAttributeValue<EntityReference>("hil_buhead");
                }
                //else if (to == "Design Head")
                //{
                //    entTo = new Entity("activityparty");
                //    entTo["partyid"] = service.Retrieve("systemuser", (regardingENtity.GetAttributeValue<EntityReference>("hil_designteam")).Id, new ColumnSet("parentsystemuserid")).GetAttributeValue<EntityReference>("parentsystemuserid");
                //}
                else if (to == "Sr. Manager Finance")
                {
                    QueryExpression query = new QueryExpression("systemuser");
                    query.ColumnSet = new ColumnSet(true);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("6532dc30-40a6-eb11-9442-6045bd72b6fd"));
                    EntityCollection _entitys = service.RetrieveMultiple(query);

                    entTo = new Entity("activityparty");
                    entTo["partyid"] = _entitys[0].ToEntityReference();
                }
                else if (to == "Sr. Manager Treasury")
                {
                    QueryExpression query = new QueryExpression("user");
                    query.ColumnSet = new ColumnSet(true);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("{45C29172-3CD0-EB11-BACC-6045BD72E9C2}"));
                    EntityCollection _entitys = service.RetrieveMultiple(query);

                    entTo = new Entity("activityparty");
                    entTo["partyid"] = _entitys[0].ToEntityReference();
                }
                entToList.Entities.Add(entTo);
            }
            return entToList;
        }
        public static void sendEmal(EntityReference approver, EntityCollection copyto, EntityReference regarding, string mailbody, string subject, IOrganizationService service)
        {
            try
            {
                Entity entEmail = new Entity("email");

                Entity entFrom = new Entity("activityparty");
                entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
                Entity[] entFromList = { entFrom };
                entEmail["from"] = entFromList;

                EntityReference to = approver;
                Entity toActivityParty = new Entity("activityparty");
                toActivityParty["partyid"] = to;
                entEmail["to"] = new Entity[] { toActivityParty };

                ;
                Entity ccActivityParty = new Entity("activityparty");
                ccActivityParty["partyid"] = copyto;
                entEmail["cc"] = copyto;

                entEmail["subject"] = subject;
                entEmail["description"] = mailbody;

                entEmail["regardingobjectid"] = regarding;

                Guid emailId = service.Create(entEmail);

                SendEmailRequest sendEmailReq = new SendEmailRequest()
                {
                    EmailId = emailId,
                    IssueSend = true
                };
                SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
        public static string createEmailBody(Entity target, string mailbodyTemp, String approver, ColumnSet collSet, IOrganizationService service, ITracingService tracingService)
        {
            tracingService.Trace("createEmailBody Started");
            try
            {
                string recordURL = URL + target.LogicalName + "&id=" + target.Id;
                string clickHere = "<br> For more information please <a href =\"" + recordURL + "\"  target=\"_blank\"><b>click here.<b></a><br>";

                Dictionary<string, string> keyValue1 = new Dictionary<string, string>();
                keyValue1.Add("approverName", approver);
                keyValue1.Add("ClickHere", clickHere);

                foreach (string spl in collSet.Columns)
                {
                    if (target.Contains(spl))
                    {
                        var dataType = target[spl].GetType();
                        if (dataType.Name == "EntityReference")
                        {
                            keyValue1.Add(spl, ((EntityReference)target[spl]).Name);
                        }
                        else if (dataType.Name.ToLower() == "string")
                        {
                            keyValue1.Add(spl, ((string)target[spl]));
                        }
                        else if (dataType.Name.ToLower() == "datetime")
                        {
                            keyValue1.Add(spl, ((DateTime)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "decial")
                        {
                            keyValue1.Add(spl, ((decimal)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "Money".ToLower())
                        {
                            keyValue1.Add(spl, ((Money)target[spl]).Value.ToString());
                        }
                        else if (dataType.Name.ToLower() == "OptionSetValue".ToLower())
                        {
                            keyValue1.Add(spl, target.FormattedValues[spl].ToString());
                        }
                        else if (dataType.Name.ToLower() == "Boolean".ToLower())
                        {
                            keyValue1.Add(spl, ((bool)target[spl]).ToString());
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException(" Error:- Data Type not found ");
                        }
                    }
                    else
                    {
                        keyValue1.Add(spl, "");
                    }
                }
                foreach (KeyValuePair<string, string> keyValue in keyValue1)
                {
                    string key = "{" + keyValue.Key + "}";
                    mailbodyTemp = mailbodyTemp.Replace(key, keyValue.Value);
                }
                // tracingService.Trace("createEmailBody Ended");
                return mailbodyTemp;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("createEmailBody Error:- " + ex.Message);
            }
        }
        public static string createEmailSubject(Entity target1, string _primaryField, string subjectTemplate, IOrganizationService service, ITracingService tracingService)
        {
            // tracingService.Trace("createEmailBody Started");
            try
            {
                Dictionary<string, string> keyValue1 = new Dictionary<string, string>();


                ColumnSet collSet = findEntityColl(subjectTemplate, _primaryField, service, tracingService);
                Entity target = service.Retrieve(target1.LogicalName, target1.Id, collSet);

                foreach (string spl in collSet.Columns)
                {
                    if (target.Contains(spl))
                    {
                        var dataType = target[spl].GetType();
                        if (dataType.Name == "EntityReference")
                        {
                            keyValue1.Add(spl, ((EntityReference)target[spl]).Name);
                        }
                        else if (dataType.Name.ToLower() == "string")
                        {
                            keyValue1.Add(spl, ((string)target[spl]));
                        }
                        else if (dataType.Name.ToLower() == "datetime")
                        {
                            keyValue1.Add(spl, ((DateTime)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "decial")
                        {
                            keyValue1.Add(spl, ((decimal)target[spl]).ToString());
                        }
                        else if (dataType.Name.ToLower() == "Money".ToLower())
                        {
                            keyValue1.Add(spl, ((Money)target[spl]).Value.ToString());
                        }
                        else if (dataType.Name.ToLower() == "OptionSetValue".ToLower())
                        {
                            keyValue1.Add(spl, target.FormattedValues[spl].ToString());
                        }
                        else if (dataType.Name.ToLower() == "Boolean".ToLower())
                        {
                            keyValue1.Add(spl, ((bool)target[spl]).ToString());
                        }
                    }
                    else
                    {
                        keyValue1.Add(spl, "");
                    }
                }
                foreach (KeyValuePair<string, string> keyValue in keyValue1)
                {
                    string key = "{" + keyValue.Key + "}";
                    subjectTemplate = subjectTemplate.Replace(key, keyValue.Value);
                }
                // tracingService.Trace("createEmailBody Ended");
                return subjectTemplate;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("createEmailSubject Error:- " + ex.Message);
            }
        }
        public static ColumnSet findEntityColl(string mailbodyTemp, string _primaryField, IOrganizationService service, ITracingService tracingService)
        {
            tracingService.Trace("findEntityColl Started");
            string[] split1 = mailbodyTemp.Split('{');
            string split2 = string.Empty;
            tracingService.Trace("1 " + split1.Length);
            foreach (string spl in split1)
            {
                if (spl.Contains("}"))
                {

                    string[] spl1 = spl.Split('}');
                    if (spl1[0] != "approverName" && spl1[0] != "ClickHere")
                        if (split2 == string.Empty)
                            split2 = split2 + spl1[0];
                        else if (!split2.Contains(spl1[0]))
                            split2 = split2 + "," + spl1[0];
                }
            }
            if (!split2.Contains("ownerid"))
                split2 = split2 + ",ownerid";

            if (!split2.Contains(_primaryField))
                split2 = split2 + _primaryField;
            string[] coll = split2.Split(',');
            ColumnSet cc = new ColumnSet(coll);
            tracingService.Trace("findEntityColl End");
            return cc;
        }
        public static void GetPrimaryIdFieldName(string _entityName, IOrganizationService service, out string _primaryField)
        {
            //Create RetrieveEntityRequest

            _primaryField = "";
            try
            {
                RetrieveEntityRequest retrievesEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = _entityName
                };
                //Execute Request
                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrievesEntityRequest);
                //_primaryKey = retrieveEntityResponse.EntityMetadata.PrimaryIdAttribute;
                _primaryField = retrieveEntityResponse.EntityMetadata.PrimaryNameAttribute;
            }
            catch (Exception ex)
            {
                _primaryField = "ERROR";
                _primaryField = ex.Message;
            }
        }
    }
}
