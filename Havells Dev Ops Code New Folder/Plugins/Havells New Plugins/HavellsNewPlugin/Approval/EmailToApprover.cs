using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.Approval
{
    public class EmailToApprover : CodeActivity
    {
        public static ITracingService tracingService = null;

        [RequiredArgument]
        [Input("Approval")]

        [ReferenceTarget("hil_approval")]
        public InArgument<EntityReference> Approval { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            try
            {

                var context = executionContext.GetExtension<IWorkflowContext>();
                var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                tracingService = executionContext.GetExtension<ITracingService>();
                EntityReference _approvalRef = this.Approval.Get((ActivityContext)executionContext);

                tracingService.Trace("1");
                Entity _approvalEntity = service.Retrieve(_approvalRef.LogicalName, _approvalRef.Id, new ColumnSet(true));

                EntityReference _regarding = _approvalEntity.GetAttributeValue<EntityReference>("regardingobjectid");
                string _primaryField = string.Empty;
                ApprovalHelper.GetPrimaryIdFieldName(_regarding.LogicalName, service, out _primaryField);
                string[] subject = _approvalEntity.GetAttributeValue<string>("subject").Split('_');
                string purpose = subject[0];
                int level = _approvalEntity.GetAttributeValue<OptionSetValue>("hil_level").Value;
                tracingService.Trace("2");
                EntityReference approver = _approvalEntity.GetAttributeValue<EntityReference>("ownerid");

                QueryExpression _query = new QueryExpression("hil_approvalmatrix");
                _query.ColumnSet = new ColumnSet("hil_approvalmatrixid", "hil_name", "createdon",
                       "statecode", "hil_purpose", "hil_mailbody", "hil_level", "hil_entity", "hil_approver",
                       "hil_approverposition", "hil_copytoposition", "hil_duehrs", "hil_kpi");
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                tracingService.Trace("_regarding.LogicalName " + _regarding.LogicalName);
                tracingService.Trace("purpose " + purpose);
                tracingService.Trace("level " + level);
                tracingService.Trace("2.9");


                _query.Criteria.AddCondition(new ConditionExpression("hil_entity", ConditionOperator.Equal, _regarding.LogicalName));
                _query.Criteria.AddCondition(new ConditionExpression("hil_purpose", ConditionOperator.Equal, purpose));
                _query.Criteria.AddCondition(new ConditionExpression("hil_level", ConditionOperator.Equal, level));
                _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                EntityCollection approvelMatrixColl = service.RetrieveMultiple(_query);
                if (approvelMatrixColl.Entities.Count <= 0)
                {
                    if (_regarding.Id != new Guid("06e1df48-690b-ee11-8f6e-6045bdac51bc"))
                        throw new InvalidPluginExecutionException("Approval Configuration Not Found..");
                    else
                    {
                        Entity _appUpdate1 = new Entity(_approvalEntity.LogicalName);
                        _appUpdate1.Id = _approvalEntity.Id;
                        _appUpdate1["hil_requesteddate"] = DateTime.Now.AddMinutes(330);
                        _appUpdate1["hil_approvalstatus"] = new OptionSetValue(3);
                        service.Update(_appUpdate1);

                        return;
                    }
                }

                tracingService.Trace("3");
                #region update Requested Date and Due Date
                Entity _appUpdate = new Entity(_approvalEntity.LogicalName);
                _appUpdate.Id = _approvalEntity.Id;
                _appUpdate["hil_requesteddate"] = DateTime.Now.AddMinutes(330);
                _appUpdate["hil_approvalstatus"] = new OptionSetValue(3);
                if (approvelMatrixColl[0].Contains("hil_duehrs"))
                {
                    _appUpdate["scheduledend"] = DateTime.Now.AddHours(approvelMatrixColl[0].GetAttributeValue<int>("hil_duehrs")).AddMinutes(330);
                }
                service.Update(_appUpdate);
                #endregion
                tracingService.Trace("3.1");
                ColumnSet collSet = ApprovalHelper.findEntityColl(approvelMatrixColl[0].GetAttributeValue<string>("hil_mailbody"), _primaryField, service, tracingService);
                Entity target = service.Retrieve(_regarding.LogicalName, _regarding.Id, collSet);
                tracingService.Trace("_primaryField " + _primaryField);
                tracingService.Trace("purpose " + purpose);
                //tracingService.Trace("target[_primaryField] " + target[_primaryField]);
                //tracingService.Trace("3.3");
                //  if(_regarding.LogicalName== "hil_tenderbankguarantee")
                {
                    //string Mailsubject = "Approval Required for " + purpose + " ID " + target[_primaryField];
                    //tracingService.Trace("Mailsubject " + Mailsubject);
                    string mailbodytemp = approvelMatrixColl[0].GetAttributeValue<string>("hil_mailbody");
                    string approverName = approver.Name;

                    tracingService.Trace("mailbodytemp " + mailbodytemp);

                    EntityReference targetOwner = target.GetAttributeValue<EntityReference>("ownerid");
                    tracingService.Trace("targetOwner " + targetOwner.Name);
                    EntityCollection entToList = new EntityCollection();
                    entToList.EntityName = "systemuser";
                    string copyto = string.Empty;
                    if (approvelMatrixColl[0].Contains("hil_copytoposition"))
                    {
                        copyto = approvelMatrixColl[0].GetAttributeValue<string>("hil_copytoposition");
                        _query = new QueryExpression("hil_userbranchmapping");
                        _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead");
                        _query.Criteria = new FilterExpression(LogicalOperator.And);
                        _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
                        _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection userMapingColl = service.RetrieveMultiple(_query);
                        if (copyto != string.Empty)
                            entToList = ApprovalHelper.getCopyToData(copyto, userMapingColl, service);

                    }
                    string Mailsubject = ApprovalHelper.createEmailSubject(target, _primaryField, approvelMatrixColl[0].GetAttributeValue<string>("hil_kpi"), service, tracingService);

                    string mailbody = ApprovalHelper.createEmailBody(target, mailbodytemp, approverName, collSet, service, tracingService);
                    tracingService.Trace("4");

                    Entity entTo = new Entity("activityparty");
                    entTo["partyid"] = targetOwner;
                    entToList.Entities.Add(entTo);
                    tracingService.Trace("5");
                    ApprovalHelper.sendEmal(approver, entToList, _regarding, mailbody, Mailsubject, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
    }
}