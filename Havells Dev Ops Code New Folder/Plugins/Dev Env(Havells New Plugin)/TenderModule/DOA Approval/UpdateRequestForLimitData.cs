using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;

namespace HavellsNewPlugin.TenderModule.DOA_Approval
{
    public class UpdateRequestForLimitData : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Contains("hil_requestforlimit"))
                {
                    QueryExpression query = new QueryExpression("hil_oaheader");
                    query.ColumnSet = new ColumnSet("hil_requestforlimit");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, entity.Id));
                    EntityCollection userMapping = service.RetrieveMultiple(query);
                    if (userMapping.Entities.Count == 1)
                    {
                        Entity oaheader = new Entity("hil_oaheader");
                        oaheader.Id = userMapping[0].Id;
                        oaheader["hil_requestforlimit"] = entity["hil_requestforlimit"];
                        service.Update(oaheader);
                    }
                }
            }
        }
    }
}
