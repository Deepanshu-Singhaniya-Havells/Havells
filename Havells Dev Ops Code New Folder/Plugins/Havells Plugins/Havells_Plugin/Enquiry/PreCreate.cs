using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.Enquiry
{
    class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "lead"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorder Job = entity.ToEntity<msdyn_workorder>();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Enquiry.PreCreate: " + ex.Message);
            }
            #endregion
        }
        public Int32 GetDailyCount(IOrganizationService service) {
            Int32 count = 1;
            try
            {
                QueryExpression qsLead = new QueryExpression("lead");
                qsLead.ColumnSet = new ColumnSet("hil_dailycount");
                EntityCollection collect_Lead = service.RetrieveMultiple(qsLead);
                qsLead.AddOrder("hil_dailycount", OrderType.Descending);
                qsLead.TopCount = 1;
                if (collect_Lead.Entities.Count > 0)
                {
                    if(collect_Lead[0].Attributes.Contains("hil_dailycount"))
                    {
                        count = collect_Lead[0].GetAttributeValue<Int32>("hil_dailycount");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Enquiry.PreCreate.GetDailyCount: " + ex.Message);
            }
            return count;
        }

    }
}
