using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class GenerateDeliverSchedule : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace(entity.LogicalName);
                    QueryExpression query = new QueryExpression("hil_deliveryschedule");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_tenderproduct", ConditionOperator.Equal, entity.Id));
                    EntityCollection deliveryCol = service.RetrieveMultiple(query);
                    if (deliveryCol.Entities.Count == 0)
                    {
                        entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_poqty", "hil_orderchecklistid"));
                        Decimal poqty = Convert.ToDecimal(entity.GetAttributeValue<decimal>("hil_poqty"));
                        Entity delSchedule = new Entity("hil_deliveryschedule");


                        Entity entityOCL = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_orderchecklistid").LogicalName, 
                            entity.GetAttributeValue<EntityReference>("hil_orderchecklistid").Id, new ColumnSet("hil_typeoforder"));
                        int typeofOder = entityOCL.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value;
                        if (typeofOder == 1)
                            delSchedule["hil_deliverydate"] = DateTime.Now; //DateTime.Now;
                        else
                            delSchedule["hil_deliverydate"] = DateTime.Now.AddDays(21);//10 to 21

                        delSchedule["hil_quantity"] = poqty;
                        delSchedule["hil_tenderproduct"] = entity.ToEntityReference();
                        service.Create(delSchedule);
                    }
                    else if (deliveryCol.Entities.Count == 1)
                    {
                        entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_poqty", "hil_orderchecklistid"));
                        Decimal poqty = Convert.ToDecimal(entity.GetAttributeValue<decimal>("hil_poqty"));
                        Entity delSchedule = new Entity("hil_deliveryschedule");
                        delSchedule.Id = deliveryCol[0].Id;

                        Entity entityOCL = service.Retrieve(entity.GetAttributeValue<EntityReference>("hil_orderchecklistid").LogicalName,
                            entity.GetAttributeValue<EntityReference>("hil_orderchecklistid").Id, new ColumnSet("hil_typeoforder"));
                        int typeofOder = entityOCL.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value;
                        if (typeofOder == 1)
                            delSchedule["hil_deliverydate"] = DateTime.Now; //DateTime.Now;
                        else
                            delSchedule["hil_deliverydate"] = DateTime.Now.AddDays(21); // 10 to 21

                        delSchedule["hil_quantity"] = poqty;
                        delSchedule["hil_tenderproduct"] = entity.ToEntityReference();
                        service.Update(delSchedule);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :" + ex.Message);
            }
        }
    }
}
