using HavellsNewPlugin.Helper;
using HavellsNewPlugin.SAP_IntegrationForOrderCreation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.RetriveMultiplePlugins
{
    public class SMSRetriveMultiple : IPlugin
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

                var inp = context.InputParameters;
                if (context.InputParameters.Contains("Query") && context.MessageName == "RetrieveMultiple")
                {
                    IntegrationConfig intConfigCreditCheck = getConfiguration(service, "SMS_Retrive");
                    string templateIDString = intConfigCreditCheck.uri;
                    string user = intConfigCreditCheck.Auth;
                    Entity userEntity = service.Retrieve("systemuser", userId, new ColumnSet("fullname"));
                    if (userEntity.GetAttributeValue<string>("fullname") != user)
                    {
                        string[] templateID = templateIDString.Split(';');
                        if (context.InputParameters["Query"] is FetchExpression)
                        {
                            tracingService.Trace("fetch");
                            FetchExpression qeft = (FetchExpression)context.InputParameters["Query"];
                            //qeft.
                            FetchXmlToQueryExpressionRequest fetchXmlToQueryExpressionRequest = new FetchXmlToQueryExpressionRequest()
                            {
                                FetchXml = qeft.Query
                            };
                            FetchXmlToQueryExpressionResponse fetchXmlToQueryExpressionResponse = (service.Execute(fetchXmlToQueryExpressionRequest) as FetchXmlToQueryExpressionResponse);
                            qe = (QueryExpression)fetchXmlToQueryExpressionResponse.Query;
                            qe.Criteria.AddCondition("hil_smstemplate", ConditionOperator.NotIn, templateID);
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
                            qe.Criteria.AddCondition("hil_smstemplate", ConditionOperator.NotIn, templateID );
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
        IntegrationConfig getConfiguration(IOrganizationService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            return output;
        }
    }
}
