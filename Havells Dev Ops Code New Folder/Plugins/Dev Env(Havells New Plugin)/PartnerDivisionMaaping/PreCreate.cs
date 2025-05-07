using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.PartnerDivisionMaaping
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
            try
            {
                OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Contains("hil_franchiseecode"))
                {
                    QueryExpression query = new QueryExpression("account");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("accountnumber", ConditionOperator.Equal, entity.GetAttributeValue<string>("hil_franchiseecode")));
                    EntityCollection entColl = service.RetrieveMultiple(query);
                    entity["hil_franchiseedirectengineer"] = entColl[0].ToEntityReference();
                }
                if (entity.Contains("hil_divisioncode"))
                {

                    QueryExpression query = new QueryExpression("product");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_sapcode", ConditionOperator.Equal, entity.GetAttributeValue<string>("hil_divisioncode"));
                    query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 2);
                    EntityCollection entityCollection = service.RetrieveMultiple((QueryBase)query);
                    if (entityCollection.Entities.Count > 0)
                        entity["hil_franchiseedirectengineer"] = entityCollection.Entities[0].ToEntityReference();
                   // GetDivisision(entity.GetAttributeValue<string>("hil_divisioncode"),service).ToEntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}
