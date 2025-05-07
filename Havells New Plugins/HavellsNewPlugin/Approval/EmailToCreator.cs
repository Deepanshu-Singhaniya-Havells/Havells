using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;
using System.Activities;

namespace HavellsNewPlugin.Approval
{

    public class EmailToCreator : CodeActivity
    {
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

                EntityReference _approvalRef = this.Approval.Get((ActivityContext)executionContext);
                Entity _approvalEntity = service.Retrieve(_approvalRef.LogicalName, _approvalRef.Id, new ColumnSet(true));

                EntityReference _regarding = _approvalEntity.GetAttributeValue<EntityReference>("regardingobjectid");


                string _primaryField = string.Empty;
                ApprovalHelper.GetPrimaryIdFieldName(_regarding.LogicalName, service, out _primaryField);

                string subject = _approvalEntity.GetAttributeValue<string>("subject") + " of " + _regarding.Name + " is " + _approvalEntity.FormattedValues["hil_approvalstatus"];
                string mailbody = "";
                bool nextIsOptional = _approvalEntity.GetAttributeValue<bool>("hil_nextisoptional");//yes
                bool sendtoOptional = _approvalEntity.GetAttributeValue<bool>("hil_sendtooptionalapproval");//no
                if (_approvalEntity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value == 1)
                {
                    if (_approvalEntity.Contains("hil_nextapproval"))
                    {
                        mailbody = @"<div data-wrapper='true' style=''><div style=''><span style='font-family:'Times New Roman',Times,serif;'><span style='font-size: 14.6667px;'>Dear User,</span></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><span style='font-size: 14.6667px;'>This is to inform you that approval applied for " + _approvalEntity.FormattedValues["hil_level"] + @" has been Approved" + (_approvalEntity.Contains("description") ? "with remarks: <b>" + _approvalEntity.GetAttributeValue<string>("description") + "</b>" : "") + @". System has requested next level Approval for further needful.</span></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>*******THIS IS A SYSTEM GENERATED NOTIFICATION.*******&nbsp;</span></strong></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>Regards,</span></strong></span></div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>System</span></strong></span></div></div>";

                        if (_regarding.LogicalName == "hil_tenderbankguarantee")
                            sendEmailtoCreator(_regarding, _approvalEntity, subject, mailbody, service);
                    }
                    else
                    {
                        mailbody = @"<div data-wrapper='true' style=''><div style=''><span style='font-family:'Times New Roman',Times,serif;'><span style='font-size: 14.6667px;'>Dear User,</span></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><span style='font-size: 14.6667px;'>This is to inform you that approval applied for and processed." + (_approvalEntity.Contains("description") ? "With remarks: <b>" + _approvalEntity.GetAttributeValue<string>("description") + "</b>" : "")
                      + @"</span></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>*******THIS IS A SYSTEM GENERATED NOTIFICATION.*******&nbsp;</span></strong></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>Regards,</span></strong></span></div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>System</span></strong></span></div></div>";
                        if (_regarding.LogicalName == "hil_tenderbankguarantee")
                            sendEmailtoCreator(_regarding, _approvalEntity, subject, mailbody, service);
                    }
                }
                else if (_approvalEntity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value == 2)
                {
                    mailbody = @"<div data-wrapper='true' style=''><div style=''><span style='font-family:'Times New Roman',Times,serif;'><span style='font-size: 14.6667px;'>Dear User,</span></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><span style='font-size: 14.6667px;'>This is to inform you that approval applied for "
                      + _approvalEntity.FormattedValues["hil_level"] + @" has been Rejected with remarks: <b>" + _approvalEntity.GetAttributeValue<string>("description") + @"</b>.  Please contact your " + _approvalEntity.FormattedValues["hil_level"] +
                      " approver for further needful.</span></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>*******THIS IS A SYSTEM GENERATED NOTIFICATION.*******&nbsp;</span></strong></span></div><div style=''>&nbsp;</div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>Regards,</span></strong></span></div><div style=''><span style='font-family:'Times New Roman',Times,serif;'><strong><span style='font-size: 14.6667px;'>System</span></strong></span></div></div>";
                    //if (_regarding.LogicalName == "hil_tenderbankguarantee")
                    sendEmailtoCreator(_regarding, _approvalEntity, subject, mailbody, service);
                }


            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
        public static void sendEmailtoCreator(EntityReference _regarding, Entity _approvalEntity, string subject, string mailbody, IOrganizationService service)
        {
            Entity entEmail = new Entity("email");

            Entity entFrom = new Entity("activityparty");
            entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
            Entity[] entFromList = { entFrom };
            entEmail["from"] = entFromList;
            Entity target = service.Retrieve(_regarding.LogicalName, _regarding.Id, new ColumnSet("ownerid"));

            EntityReference to = target.GetAttributeValue<EntityReference>("ownerid");
            Entity toActivityParty = new Entity("activityparty");
            toActivityParty["partyid"] = to;
            entEmail["to"] = new Entity[] { toActivityParty };

            EntityCollection entCCList = new EntityCollection();
            Entity entCC1 = new Entity("activityparty");
            entCC1["partyid"] = _approvalEntity.GetAttributeValue<EntityReference>("ownerid");
            entCCList.Entities.Add(entCC1);
            if (_regarding.LogicalName == "hil_orderchecklist")
            {
                string[] subject11 = _approvalEntity.GetAttributeValue<string>("subject").Split('_');
                string purpose = subject11[0];
                int level = _approvalEntity.GetAttributeValue<OptionSetValue>("hil_level").Value;
                QueryExpression _query = new QueryExpression("hil_approvalmatrix");
                _query.ColumnSet = new ColumnSet("hil_approvalmatrixid", "hil_name", "createdon",
                       "statecode", "hil_purpose", "hil_mailbody", "hil_level", "hil_entity", "hil_approver",
                       "hil_approverposition", "hil_copytoposition", "hil_duehrs", "hil_kpi");
                _query.Criteria = new FilterExpression(LogicalOperator.And);

                _query.Criteria.AddCondition(new ConditionExpression("hil_entity", ConditionOperator.Equal, _regarding.LogicalName));
                _query.Criteria.AddCondition(new ConditionExpression("hil_purpose", ConditionOperator.Equal, purpose));
                _query.Criteria.AddCondition(new ConditionExpression("hil_level", ConditionOperator.Equal, level));
                _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                EntityCollection approvelMatrixColl = service.RetrieveMultiple(_query);
                string copyto = string.Empty;
                if (approvelMatrixColl[0].Contains("hil_copytoposition"))
                {
                    copyto = approvelMatrixColl[0].GetAttributeValue<string>("hil_copytoposition");
                    _query = new QueryExpression("hil_userbranchmapping");
                    _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead");
                    _query.Criteria = new FilterExpression(LogicalOperator.And);
                    _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, to.Id));
                    _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection userMapingColl = service.RetrieveMultiple(_query);
                    if (copyto != string.Empty)
                        entCCList = ApprovalHelper.getCopyToData(copyto, userMapingColl, service);
                }

            }
            entCCList.Entities.Add(entCC1);
            entEmail["cc"] = entCCList;


            entEmail["subject"] = subject;
            entEmail["description"] = mailbody;

            entEmail["regardingobjectid"] = _regarding;

            Guid emailId = service.Create(entEmail);

            SendEmailRequest sendEmailReq = new SendEmailRequest()
            {
                EmailId = emailId,
                IssueSend = true
            };
            SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);
        }
    }
}
