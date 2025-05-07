using System;
using HavellsNewPlugin.Helper;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.RetriveMultiplePlugins
{
    class RestrictMaster : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            QueryExpression qe = null;
            try
            {
                Guid userId = context.UserId;
                if (HelperClass.getUserSecurityRole(userId, service, "System Administrator", tracingService))
                {
                    var inp = context.InputParameters;
                    if (context.InputParameters.Contains("Query") && context.MessageName == "RetrieveMultiple")
                    {
                        Entity user = service.Retrieve("systemuser", userId, new ColumnSet("hil_department"));
                        Guid department = user.GetAttributeValue<EntityReference>("hil_department").Id;
                        if (context.InputParameters["Query"] is FetchExpression)
                        {
                            tracingService.Trace("fetch");
                            FetchExpression qeft = (FetchExpression)context.InputParameters["Query"];
                            FetchXmlToQueryExpressionRequest fetchXmlToQueryExpressionRequest = new FetchXmlToQueryExpressionRequest()
                            {
                                FetchXml = qeft.Query
                            };
                            FetchXmlToQueryExpressionResponse fetchXmlToQueryExpressionResponse = (service.Execute(fetchXmlToQueryExpressionRequest) as FetchXmlToQueryExpressionResponse);
                            qe = (QueryExpression)fetchXmlToQueryExpressionResponse.Query;
                            qe.Criteria.AddCondition("hil_department", ConditionOperator.Equal, department);
                            QueryExpressionToFetchXmlRequest queryExpressionToFetchXmlRequest = new QueryExpressionToFetchXmlRequest()
                            {
                                Query = qe
                            };
                            QueryExpressionToFetchXmlResponse queryExpressionToFetchXmlResponse = (service.Execute(queryExpressionToFetchXmlRequest) as QueryExpressionToFetchXmlResponse);
                            context.InputParameters["Query"] = new FetchExpression(queryExpressionToFetchXmlResponse.FetchXml);
                        }
                        else if (context.InputParameters["Query"] is QueryExpression)
                        {
                            tracingService.Trace("query");
                            qe = (QueryExpression)context.InputParameters["Query"];
                            qe.Criteria.AddCondition("hil_department", ConditionOperator.Equal, department);
                            context.InputParameters["Query"] = qe;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}
