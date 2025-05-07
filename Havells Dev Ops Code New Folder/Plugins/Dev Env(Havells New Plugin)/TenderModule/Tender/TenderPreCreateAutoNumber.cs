using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.Tender
{
   public class TenderPreCreateAutoNumber : IPlugin
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
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "hil_tender")
            {
                tracingService.Trace("Target");
                OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                Entity entity = (Entity)context.InputParameters["Target"];
                //  Entity _entTender = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_department", "hil_name"));
                Guid department = entity.GetAttributeValue<EntityReference>("hil_department").Id;
                tracingService.Trace("department " + department);
                tracingService.Trace("entity.LogicalName " + entity.LogicalName);
                QueryExpression QueryIdg = new QueryExpression("plt_idgenerator");
                QueryIdg.ColumnSet = new ColumnSet(true);
                QueryIdg.Criteria = new FilterExpression(LogicalOperator.And);
                QueryIdg.Criteria.AddCondition("hil_department", ConditionOperator.Equal, department);
                QueryIdg.Criteria.AddCondition("plt_idgen_name", ConditionOperator.Equal, entity.LogicalName);
                EntityCollection Found = service.RetrieveMultiple(QueryIdg);
                string AttributeName = "";
                int FixedNumberSize = 0;
                int NextNumber = 0;
                int IncrementBy = 0;
                string Prefix = "";
                if (Found.Entities.Count > 0)
                {
                    tracingService.Trace("2");
                    foreach (Entity ent in Found.Entities)
                    {
                        AttributeName = ent.GetAttributeValue<string>("plt_attributename").ToString();
                        FixedNumberSize = ent.GetAttributeValue<int>("plt_idgen_fixednumbersize");
                        NextNumber = ent.GetAttributeValue<int>("plt_idgen_nextnumber");
                        IncrementBy = ent.GetAttributeValue<int>("plt_idgen_incrementby");
                        Prefix = ent.GetAttributeValue<string>("plt_idgen_prefix");
                        tracingService.Trace("AttributeName " + AttributeName + " FixedNumberSize " + FixedNumberSize.ToString() + "NextNumber " + NextNumber.ToString() + "IncrementBy " + IncrementBy.ToString() + "Prefix " + Prefix);
                    }

                }
                string nextNo = NextNumber.ToString().PadLeft(8, '0');
                tracingService.Trace("nextNo " + nextNo);
                string FinalNumber = Prefix + nextNo;
                tracingService.Trace("FinalNumber" + FinalNumber);
                int NextNum = NextNumber + IncrementBy;
                tracingService.Trace("NextNum " + NextNum);
                Entity idg = new Entity("plt_idgenerator");
                idg.Id = Found.Entities[0].Id;
                idg["plt_idgen_nextnumber"] = NextNum;
                service.Update(idg);

                entity["hil_name"] = FinalNumber.ToString();

                tracingService.Trace("Complete");

            }
        }
    }
}
