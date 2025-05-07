using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Campaign
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                bool campaignRunBy = entity.GetAttributeValue<bool>("hil_campaignrunby");
                OptionSetValue campaignMedium = entity.GetAttributeValue<OptionSetValue>("hil_campaignmedium");
                OptionSetValue campaignBrand = entity.GetAttributeValue<OptionSetValue>("hil_brand");
                string campaignCode = entity.GetAttributeValue<string>("codename");
                EntityReference campaignContent = entity.GetAttributeValue<EntityReference>("hil_campaigncontent");

                if (!campaignRunBy)// Campaign Run By Own, Generate Campaign Code automatically
                {
                    campaignCode = GenerateCampaignId(service);
                    entity["codename"] = campaignCode;
                }
                if (campaignMedium.Value != 3 && campaignMedium.Value != 4)// Generate Campaign URL is Campaign Medium is not equal to Email/SMS
                {
                    QueryExpression qsUser = new QueryExpression("hil_campaignwebsitesetup");
                    qsUser.ColumnSet = new ColumnSet("hil_baseurl");
                    ConditionExpression condExp = new ConditionExpression("hil_brand", ConditionOperator.Equal, campaignBrand.Value);
                    qsUser.Criteria.AddCondition(condExp);
                    qsUser.NoLock = true;
                    EntityCollection collect_user = service.RetrieveMultiple(qsUser);
                    if (collect_user.Entities.Count > 0)
                    {
                        entity["hil_campaignbaseurl"] = collect_user.Entities[0].GetAttributeValue<string>("hil_baseurl");
                    }
                    else {
                        throw new InvalidPluginExecutionException("Campaign Brand Setup is missing. Please contact to System Administrator");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Campaign.PreCreate.Execute  ***" + ex.Message + "***  ");
            }
        }
        private string GenerateCampaignId(IOrganizationService service)
        {
            string _retValue = string.Empty;
            int campaignCount = 1;
            try
            {
                QueryExpression qsUser = new QueryExpression("campaign");
                qsUser.ColumnSet = new ColumnSet("codename");
                ConditionExpression condExp = new ConditionExpression("hil_campaignrunby", ConditionOperator.NotEqual, true);
                qsUser.Criteria.AddCondition(condExp);
                qsUser.NoLock = true;
                EntityCollection collect_user = service.RetrieveMultiple(qsUser);
                if (collect_user.Entities.Count > 0)
                {
                    campaignCount = collect_user.Entities.Count + 1;
                }
                _retValue = "CAMP" + DateTime.Now.Year.ToString().Substring(2, 2) + campaignCount.ToString().PadLeft(5, '0');
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Campaign.PreCreate.GenerateCampaignId * **" + ex.Message + "***  ");
            }
            return _retValue;
        }
    }
}
