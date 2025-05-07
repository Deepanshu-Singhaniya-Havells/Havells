using System;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class CreateRMCostSheetSubGrid : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.InputParameters.Contains("LineId") && context.InputParameters["LineId"] is string)
                {
                    tracingService.Trace("Step-1");
                    var lineIdsString = context.InputParameters["LineId"].ToString();
                    string[] arrProductIds = lineIdsString.Split(';');
                    int id = 0; int rmsc_Count = 0;
                    tracingService.Trace(JsonConvert.SerializeObject(arrProductIds));
                    foreach (var productIdString in arrProductIds)
                    {
                        id = id++;
                        tracingService.Trace("Step-2");
                        Guid productId = new Guid(productIdString);
                        QueryExpression query = new QueryExpression("hil_tenderproduct");
                        query.ColumnSet = new ColumnSet("hil_selectproduct", "hil_tenderid", "hil_name");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, productId));
                        EntityCollection entColl = service.RetrieveMultiple(query);

                        if (entColl.Entities.Count > 0)
                        {
                            tracingService.Trace("Step-3");
                            bool foundSelectedProduct = false;
                            bool selectProduct = false;

                            foreach (Entity ent in entColl.Entities)
                            {

                                selectProduct = ent.GetAttributeValue<bool>("hil_selectproduct");
                                if (selectProduct)
                                {

                                    foundSelectedProduct = true; // Write-In Product Code
                                    //break;
                                }
                            }
                            tracingService.Trace("Step-4");
                            if (foundSelectedProduct)
                            {
                                tracingService.Trace("Step-4A or if case");
                                string TenderproductName = service.Retrieve("hil_tenderproduct", productId, new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
                                Entity entity = new Entity("hil_rmcostsheet");
                                entity["hil_name"] = $"RM{TenderproductName}";
                                EntityReference TenderProductRef = new EntityReference("hil_tenderproduct", productId);
                                entity["hil_tenderproductid"] = TenderProductRef;
                                Entity EntityTender = service.Retrieve("hil_tenderproduct", productId, new ColumnSet("hil_tenderid"));
                                Guid TenderId = EntityTender.Contains("hil_tenderid") ? EntityTender.GetAttributeValue<EntityReference>("hil_tenderid").Id : Guid.Empty;
                                if (TenderId != Guid.Empty)
                                {
                                    entity["hil_tenderid"] = new EntityReference("hil_tender", TenderId);
                                }
                                tracingService.Trace("repeat step");
                                query = new QueryExpression("hil_rmcostsheet");
                                query.ColumnSet = new ColumnSet(true);
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, productId));
                                query.Criteria.AddCondition(new ConditionExpression("hil_tenderid", ConditionOperator.Equal, TenderId));
                                entColl = service.RetrieveMultiple(query);
                                tracingService.Trace("before checking repeat case");
                                if (entColl.Entities.Count <= 0)
                                {
                                    Guid rmGuid = service.Create(entity);
                                    Entity TenderProductentity = new Entity("hil_tenderproduct", productId);
                                    TenderProductentity["hil_rmcostsheet"] = new EntityReference("hil_rmcostsheet", rmGuid);
                                    context.OutputParameters["Message"] = "RM Cost Sheets created successfully.";
                                    service.Update(TenderProductentity);
                                    rmsc_Count = rmsc_Count++;

                                }
                                if (entColl.Entities.Count > 0 && arrProductIds.Length == id && rmsc_Count == 0)
                                {
                                    tracingService.Trace("Inside");
                                    context.OutputParameters["Message"] = "RM Cost Sheet Already Created";
                                    context.OutputParameters["Status"] = true;
                                    break;
                                }
                                else if (rmsc_Count > 0 && arrProductIds.Length == id)
                                {
                                    context.OutputParameters["Message"] = "RM Cost Sheets Created Successfully";
                                    context.OutputParameters["Status"] = true;
                                }

                            }
                            else
                            {

                                tracingService.Trace("5 or else case");
                                context.OutputParameters["Message"] = "Kindly select only Write-In case to create RM Cost Sheet";
                                context.OutputParameters["Status"] = false;
                                tracingService.Trace("step - 6");
                                break;
                            }

                            if (context.MessageName.ToLower() == "delete")
                            {
                                productId = ((Entity)context.InputParameters["Target"]).Id;

                                query = new QueryExpression("hil_rmcostsheet");
                                query.ColumnSet = new ColumnSet(true);
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, productId));

                                entColl = service.RetrieveMultiple(query);

                                if (entColl.Entities.Any())
                                {
                                    foreach (Entity ent in entColl.Entities)
                                    {
                                        service.Delete(ent.LogicalName, ent.Id);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace("FaultException occurred: {0}", ex.Detail.Message);
                throw;
            }
            catch (Exception ex)
            {
                tracingService.Trace("General Exception occurred: {0}", ex.Message);
                throw;
            }
        }
    }
}
