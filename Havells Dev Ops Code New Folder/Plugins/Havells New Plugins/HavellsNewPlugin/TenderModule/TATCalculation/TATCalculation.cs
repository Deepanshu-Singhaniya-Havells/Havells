using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TATCalculation
{
    public class TATCalculation : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                int fromStage = 0;
                int toStage = 0;
                EntityReference regardingRef = new EntityReference();
                EntityReference department = new EntityReference();
                EntityReference owner = new EntityReference();
                Entity preImage = new Entity();
                Entity postImage = new Entity();

                #region PluginConfig
                tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                #endregion
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity targetEntity = (Entity)context.InputParameters["Target"];
                    //targetEntity = service.Retrieve(targetEntity.LogicalName, targetEntity.Id, new ColumnSet("hil_stakeholder", "hil_department"));
                    tracingService.Trace("Target");
                    string stackHolderField = "";
                    if (targetEntity.LogicalName == "hil_tender")
                    {
                        stackHolderField = "hil_stakeholder";
                    }
                    else if (targetEntity.LogicalName == "hil_orderchecklist")
                    {
                        stackHolderField = "hil_stackholder";
                    }
                    else if (targetEntity.LogicalName == "hil_oaheader")
                    {
                        stackHolderField = "hil_stackholder";
                    }


                    if (targetEntity.Contains(stackHolderField))
                    {
                        preImage = ((Entity)context.PreEntityImages["PreImage"]);
                        tracingService.Trace("PreImage");
                        postImage = ((Entity)context.PostEntityImages["PostImage"]);
                        tracingService.Trace("PostImage");

                        fromStage = preImage.GetAttributeValue<OptionSetValue>(stackHolderField).Value;
                        tracingService.Trace("from Stage " + fromStage);
                        toStage = postImage.GetAttributeValue<OptionSetValue>(stackHolderField).Value;
                        tracingService.Trace("toStage Stage " + toStage);
                        regardingRef = targetEntity.ToEntityReference();
                        tracingService.Trace("Ref " + regardingRef.Id);
                        department = postImage.GetAttributeValue<EntityReference>("hil_department");
                        tracingService.Trace("department " + department.Id);

                        CloseExistingTatLine(fromStage, toStage, regardingRef, service);
                        tracingService.Trace("Close Existing ");


                        if (targetEntity.LogicalName == "hil_tender")
                            createTATLine(fromStage, toStage, regardingRef, department, service, postImage.GetAttributeValue<bool>("hil_productrequisition"), "Product Code Requisition TAT");
                        else
                            createTATLine(fromStage, toStage, regardingRef, department, service, false, "Product Code Requisition TAT");
                        tracingService.Trace("create TAT Line");

                        if (toStage == 3)
                        {
                            owner = postImage.GetAttributeValue<EntityReference>("hil_designteam");
                            internalTransfer(regardingRef, service, owner);
                        }
                        if (targetEntity.LogicalName == "hil_tender")
                        {
                            if (fromStage == 3)
                            {
                                Entity entity = new Entity(targetEntity.LogicalName, targetEntity.Id);
                                entity["hil_productrequisition"] = false;
                                service.Update(entity);
                            }
                        }
                    }
                    else if (targetEntity.Contains("hil_designteam") && (targetEntity.LogicalName == "hil_tender"))
                    {
                        regardingRef = targetEntity.ToEntityReference();
                        owner = targetEntity.GetAttributeValue<EntityReference>("hil_designteam");
                        internalTransfer(regardingRef, service, owner);
                    }

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        void internalTransfer(EntityReference regardingRef, IOrganizationService service, EntityReference owner)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_bdtatownership");
                query.ColumnSet = new ColumnSet("actualstart");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, regardingRef.Id);
                query.Criteria.AddCondition("statuscode", ConditionOperator.NotIn, new object[] { 2, 3 });
                EntityCollection Found = service.RetrieveMultiple(query);
                if (Found.Entities.Count > 0)
                {
                    Entity entity = new Entity("hil_bdtatdetails");
                    entity["hil_activity"] = Found[0].ToEntityReference();
                    Guid guid = service.Create(entity);
                    Entity entity1 = new Entity(entity.LogicalName);
                    entity1.Id = guid;
                    entity1["ownerid"] = owner;
                    service.Update(entity1);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error on Internal transfer" + ex.Message);
            }
        }
        void CloseExistingTatLine(int fromStage, int toStage, EntityReference regardingRef, IOrganizationService service)
        {
            try
            {
                QueryExpression query = new QueryExpression("hil_bdtatownership");
                query.ColumnSet = new ColumnSet("actualstart");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, regardingRef.Id);
                //query.Criteria.AddCondition("hil_fromstage", ConditionOperator.Equal, new OptionSetValue(fromStage));
                //query.Criteria.AddCondition("hil_tostage", ConditionOperator.Equal, new OptionSetValue(toStage));
                query.Criteria.AddCondition("statuscode", ConditionOperator.NotIn, new object[] { 2, 3 });
                EntityCollection Found = service.RetrieveMultiple(query);
                if (Found.Entities.Count > 0)
                {
                    DateTime start = Found[0].GetAttributeValue<DateTime>("actualstart");
                    DateTime now = DateTime.Now.ToUniversalTime();
                    TimeSpan ts = now - start;
                    int duration = Convert.ToInt32(ts.TotalMinutes);
                    Entity entity = new Entity(Found[0].LogicalName, Found[0].Id);
                    entity["actualend"] = now;
                    entity["actualdurationminutes"] = duration;
                    service.Update(entity);
                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = entity.Id,
                            LogicalName = entity.LogicalName,
                        },
                        State = new OptionSetValue(1),
                        Status = new OptionSetValue(2)
                    };
                    service.Execute(setStateRequest);
                }
            }
            catch (Exception ex)
            {

                throw new InvalidPluginExecutionException("Error on Close Existing TAT Line " + ex.Message);
            }
        }
        void createTATLine(int fromStage, int toStage, EntityReference regardingRef, EntityReference department, IOrganizationService service, bool productCode, string Tatname)
        {
            try
            {

                QueryExpression query = new QueryExpression("hil_salestatmaster");
                query.ColumnSet = new ColumnSet("hil_durationmin", "hil_name", "hil_department");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_department", ConditionOperator.Equal, department.Id);

                if (toStage == 3 && productCode)
                    query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Tatname);
                else
                {
                    query.Criteria.AddCondition("hil_fromstage", ConditionOperator.Equal, fromStage);
                    query.Criteria.AddCondition("hil_tostage", ConditionOperator.Equal, toStage);
                }
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection Found = service.RetrieveMultiple(query);
                if (Found.Entities.Count > 0)
                {
                    int durationmin = Found[0].Contains("hil_durationmin") ? Found[0].GetAttributeValue<int>("hil_durationmin") : throw new InvalidPluginExecutionException("***** Duration Not Defin for this type *****"); ;
                    DateTime startTime = DateTime.Now.ToUniversalTime();
                    DateTime endTime = DateTime.Now.ToUniversalTime().AddMinutes(durationmin);
                    Entity entity = new Entity("hil_bdtatownership");
                    entity["subject"] = Found[0].Contains("hil_name") ? Found[0].GetAttributeValue<string>("hil_name") : throw new InvalidPluginExecutionException("***** name Not Defin for this type *****"); ;
                    entity["hil_fromstage"] = new OptionSetValue(fromStage);
                    entity["hil_tostage"] = new OptionSetValue(toStage);
                    entity["scheduleddurationminutes"] = durationmin;
                    entity["scheduledstart"] = startTime;
                    entity["scheduledend"] = endTime;
                    entity["hil_salestatmaster"] = Found[0].ToEntityReference();
                    entity["actualstart"] = startTime;
                    entity["regardingobjectid"] = regardingRef;
                    entity["hil_department"] = Found[0].GetAttributeValue<EntityReference>("hil_department");
                    Guid guid = service.Create(entity);
                }
                else
                {
                    //throw new InvalidPluginExecutionException("***** Master Not Defin for this type *****");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error on Create TAT Line " + ex.Message);
            }
        }
    }
}
