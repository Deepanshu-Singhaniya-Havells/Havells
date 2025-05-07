using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.ClaimLine
{
    public class PreCreate : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_claimline"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    //PopulateData(entity, service);
                    RestictDuplicateClaimLine(service, entity);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimLine.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }

        public static Int32 GetMobileAppClosurePenalt(IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    Int32 Value = 0;
                    QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                    qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ClaimJobMobileAppClosureIncentive");
                    qe.Criteria.AddCondition("hil_mobileappclosure", ConditionOperator.Equal, 1);
                    EntityCollection enColl = service.RetrieveMultiple(qe);
                    foreach (Entity en in enColl.Entities)
                    {
                        if (en.Contains("hil_priceforincentive"))
                        {
                            Value = en.GetAttributeValue<Int32>("hil_priceforincentive");
                        }
                    }
                    return Value;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimLine.PreCreate.PopulateData" + ex.Message);
            }
        }
        protected static void RestictDuplicateClaimLine(IOrganizationService service, Entity claimline)
        {
            if (claimline.Contains("hil_jobid"))
            {
                EntityReference job = claimline.GetAttributeValue<EntityReference>("hil_jobid");
                EntityReference hil_claimcategory = claimline.GetAttributeValue<EntityReference>("hil_claimcategory");
                QueryExpression qryExp = new QueryExpression("hil_claimline");
                qryExp.ColumnSet = new ColumnSet(false);
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, job.Id);
                qryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, hil_claimcategory.Id);
                EntityCollection entColClaim = service.RetrieveMultiple(qryExp);
                if (entColClaim.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Job already Closed!!! Duplicate Claim Lines are not allowed.");
                }
                else
                {
                    throw new InvalidPluginExecutionException(entColClaim.Entities.Count.ToString());
                }
            }
            else
            {
                throw new InvalidPluginExecutionException("Job Id is null");
            }
        }
    }
}
