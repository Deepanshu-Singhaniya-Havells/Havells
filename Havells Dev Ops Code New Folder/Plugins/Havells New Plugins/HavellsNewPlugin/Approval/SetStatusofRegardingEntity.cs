using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;
using System.Activities;
using HavellsNewPlugin.TenderModule.OrderCheckList;

namespace HavellsNewPlugin.Approval
{
    public class SetStatusofRegardingEntity : CodeActivity
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
                Entity _approvalEntity = service.Retrieve(_approvalRef.LogicalName, _approvalRef.Id,
                    new ColumnSet("regardingobjectid", "hil_approvalstatus", "hil_nextapproval", "hil_level"));
                EntityReference _regarding = _approvalEntity.GetAttributeValue<EntityReference>("regardingobjectid");
                if (!_approvalEntity.Contains("hil_nextapproval"))
                {
                    // throw new InvalidPluginExecutionException("withoutNext Approval");
                    Entity regardingEntity = new Entity(_regarding.LogicalName);
                    regardingEntity.Id = _regarding.Id;
                    regardingEntity["hil_approvalstatus"] = _approvalEntity["hil_approvalstatus"];
                    service.Update(regardingEntity);
                }
                else
                {

                    if (_approvalEntity.Contains("hil_level"))
                    {

                        int approvalStatus = _approvalEntity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value;
                        if (approvalStatus == 1)
                        {
                            int level = _approvalEntity.GetAttributeValue<OptionSetValue>("hil_level").Value;
                            if (level == 1)
                            {
                                if (_regarding.LogicalName == "hil_orderchecklist")
                                    SubmitForApproval.validateLP(service, service.Retrieve(_regarding.LogicalName, _regarding.Id, new ColumnSet(true)));
                                approvalStatus = 5;
                            }
                            else if (level == 2)
                            {
                                approvalStatus = 6;
                            }
                            else if (level == 3)
                            {
                                approvalStatus = 7;
                            }
                            else if (level == 4)
                            {
                                approvalStatus = 8;
                            }
                            else if (level == 5)
                            {
                                approvalStatus = 9;
                            }
                        }
                        Entity regardingEntity = new Entity(_regarding.LogicalName);
                        regardingEntity.Id = _regarding.Id;
                        regardingEntity["hil_approvalstatus"] = new OptionSetValue(approvalStatus);
                        service.Update(regardingEntity);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
    }
}
