using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;

namespace Havells_Plugin.PerformaInvoice
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            //try
            //{
            if (context.InputParameters.Contains("Target")
                && context.InputParameters["Target"] is Entity
                && context.MessageName.ToUpper() == "UPDATE" && context.PrimaryEntityName == "hil_claimheader")
            {
                OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                Entity entity = (Entity)context.InputParameters["Target"];
                ClaimOperations clmOps = new ClaimOperations();
                if (entity.Contains("hil_performastatus"))
                {
                    int _performaInvoiceStatus = entity.GetAttributeValue<OptionSetValue>("hil_performastatus").Value;
                    if (_performaInvoiceStatus == 2 || _performaInvoiceStatus == 3)
                    {
                        string _fetxhXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_claimoverheadline'>
                                <attribute name='hil_claimoverheadlineid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_productcategory' />
                                <attribute name='hil_callsubtype' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_performainvoice' operator='eq' value='{entity.Id}' />
                                  <filter type='or'>
                                    <condition attribute='hil_productcategory' operator='null' />
                                    <condition attribute='hil_callsubtype' operator='null' />
                                  </filter>
                                  <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                              </entity>
                            </fetch>";

                        EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(_fetxhXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException("There are some claim overheads which do not have Product Category/Call Subtype.");
                        }
                    }
                    if (_performaInvoiceStatus == 2) // Submit for approval
                    {
                        clmOps.GenerateFixedCompensationLines(service, entity.Id);
                        clmOps.UpdatePerformaInvoice(service, entity.Id);
                    }
                    else if (_performaInvoiceStatus == 3) // Approve
                    {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_claimheader'>
                                <attribute name='hil_claimheaderid' />
                                <filter type='and'>
                                    <condition attribute='hil_claimheaderid' operator='eq' value='{entity.Id}' />
                                </filter>
                                <link-entity name='account' from='accountid' to='hil_franchisee' link-type='inner' alias='cp'>
                                    <attribute name='hil_salesoffice' />
                                    <attribute name='hil_vendorcode' />
                                </link-entity>
                                </entity>
                                </fetch>";
                        EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (!entcoll.Entities[0].Contains("cp.hil_salesoffice") || !entcoll.Entities[0].Contains("cp.hil_vendorcode"))
                        {
                            throw new InvalidPluginExecutionException("Sales Office|Vendor Code is missing at Channel Partner Master record. Please Contact to Service CRM Admin.");
                        }

                        QueryExpression queryExp = new QueryExpression("hil_claimline");
                        queryExp.ColumnSet = new ColumnSet(false);
                        queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExp.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, entity.Id);
                        queryExp.Criteria.AddCondition("hil_activitycode", ConditionOperator.Null);
                        queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

                        entcoll = service.RetrieveMultiple(queryExp);
                        if (entcoll.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException("There are some claim lines which do not have Activity Code mapped with. Please Contact to Service CRM Admin.");
                        }
                        else
                        {
                            clmOps.GenerateClaimOverHeads(service, entity.Id);
                            clmOps.UpdatePerformaInvoice(service, entity.Id);
                            clmOps.GenerateClaimSummary(service, entity.Id);
                        }
                    }

                    //clmOps.UpdatePerformaInvoice(service, entity.Id);
                }
            }
            //}
            //catch (Exception ex)
            //{
            //    throw new InvalidPluginExecutionException("  ***Havells_Plugin.PerformaInvoice.PostUpdate.Execute***  " + ex.Message);
            //}
        }
    }
}
