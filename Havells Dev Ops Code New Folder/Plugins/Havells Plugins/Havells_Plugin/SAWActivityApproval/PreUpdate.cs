using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace Havells_Plugin.SAWActivityApproval
{
    public class PreUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                Entity preImage = (Entity)context.PreEntityImages["image"];
                OptionSetValue optVal = null;
                EntityReference erSAWActivity = null;

                if (preImage.Contains("hil_approver"))
                {
                    EntityReference erApprover = preImage.GetAttributeValue<EntityReference>("hil_approver");
                    if (erApprover.Id != context.UserId) {
                        throw new InvalidPluginExecutionException(" ***You are not authorized to Approve at this Level*** ");
                    }
                }

                if (preImage.Contains("hil_sawactivity"))
                {
                    erSAWActivity = preImage.GetAttributeValue<EntityReference>("hil_sawactivity");
                }
                if (preImage.Contains("hil_level") && preImage.Contains("hil_sawactivity"))
                {
                    optVal = preImage.GetAttributeValue<OptionSetValue>("hil_level");

                    if (optVal.Value != 1)
                    {
                        QueryExpression qrExp;
                        EntityCollection entCol;

                        qrExp = new QueryExpression("hil_sawactivityapproval");
                        qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid", "hil_approvalstatus");
                        qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qrExp.Criteria.AddCondition("hil_sawactivity", ConditionOperator.Equal, erSAWActivity.Id);
                        qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.NotEqual, 3);
                        if (optVal.Value == 2)
                        {
                            qrExp.Criteria.AddCondition("hil_level", ConditionOperator.In, new object[] { 1 });
                        }
                        else if (optVal.Value == 3)
                        {
                            qrExp.Criteria.AddCondition("hil_level", ConditionOperator.In, new object[] { 2 });
                        }
                        entCol = service.RetrieveMultiple(qrExp);
                        if (entCol.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException(" ***SAW Approval Is pending At Previous Level*** ");
                        }
                    }
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SAWActivityApproval.PreUpdate" + e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SAWActivityApproval.PreUpdate" + e.Message);
            }
        }
    }
}
