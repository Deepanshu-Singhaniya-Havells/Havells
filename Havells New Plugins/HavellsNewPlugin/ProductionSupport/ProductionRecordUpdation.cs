using System;
using HavellsNewPlugin.Helper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.ProductionSupport
{
    public class ProductionRecordUpdation : IPlugin
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
            #region GetField
            try
            {
                if (context.InputParameters.Contains("Target")
                && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_production"
                && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string _rowIds = string.Empty;
                    Guid user = context.UserId;
                    tracingService.Trace("user " + user);
                    if (context.MessageName.ToUpper() == "CREATE")
                    {
                        string SchemaName = entity.GetAttributeValue<string>("hil_schemaname").ToString();
                        string RecordType = entity.FormattedValues["hil_recordtype"].ToString();
                        string newvalue = entity.GetAttributeValue<string>("hil_newvalue");
                        EntityReference entityrecord = null;
                        if (entity.Contains("hil_tenderno"))
                        {
                            entityrecord = entity.GetAttributeValue<EntityReference>("hil_tenderno");
                        }
                        else if (entity.Contains("hil_orderchecklist"))
                        {
                            entityrecord = entity.GetAttributeValue<EntityReference>("hil_orderchecklist");
                        }
                        tracingService.Trace("entityrecord.LogicalName " + entityrecord.LogicalName);
                        tracingService.Trace("RecordType " + RecordType);
                        tracingService.Trace("SchemaName " + SchemaName);

                        RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = entityrecord.LogicalName,
                            LogicalName = SchemaName,
                            RetrieveAsIfPublished = false
                        };
                        RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);

                        tracingService.Trace("Retrieved the attribute {0}.", attributeResponse.AttributeMetadata.SchemaName);
                        var attri = attributeResponse.AttributeMetadata;
                        if (attri.AttributeType == AttributeTypeCode.Picklist)
                        {
                            string optionLable = string.Empty;
                            var option = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)attri).OptionSet.Options;
                            int value = -1;
                            for (int j = 0; j < option.Count; j++)
                            {
                                optionLable += option[j].Label.LocalizedLabels[0].Label + " ,";
                                //throw new InvalidPluginExecutionException("Convert.ToInt32(option[j].Value); " + Convert.ToInt32(option[j].Value));
                                if (option[j].Label.LocalizedLabels[0].Label.ToString().ToLower() == newvalue.ToLower())
                                {
                                    value = Convert.ToInt32(option[j].Value);
                                    //throw new InvalidPluginExecutionException("newvalue " + newvalue + " value " + value);
                                }
                            }
                            if (value == -1)
                            {
                                throw new InvalidPluginExecutionException("Please enter correct value out of these " + optionLable);
                            }
                        }
                        else if (attri.AttributeType == AttributeTypeCode.Lookup)
                        {
                            //throw new InvalidPluginExecutionException("newvalue" + newvalue);
                            var TargetEntity = ((Microsoft.Xrm.Sdk.Metadata.LookupAttributeMetadata)attri).Targets;
                            //throw new InvalidPluginExecutionException("TargetEntity" + TargetEntity[0].ToString());
                             QueryExpression targetLook = new QueryExpression(TargetEntity[0].ToString());
                            targetLook.ColumnSet = new ColumnSet(false);
                            targetLook.Criteria = new FilterExpression(LogicalOperator.And);
                            targetLook.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, newvalue);
                            EntityCollection entity1 = service.RetrieveMultiple(targetLook);

                            if (entity1.Entities.Count == 0)
                            {
                                throw new InvalidPluginExecutionException("Please enter correct Customer Code  ");
                            }
                        }
                        SetApprover(service, entity, tracingService);
                    }
                    else
                    {
                        int approvalStatus = Convert.ToInt32(entity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value);
                        tracingService.Trace("approvalStatus" + approvalStatus);
                        if (approvalStatus == 1)
                        {
                            entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_approvalstatus", "hil_recordtype", "hil_tenderno", "hil_orderchecklist", "hil_selectfield", "hil_oldvalue", "hil_displayname", "hil_newvalue", "hil_datatype", "hil_schemaname", "ownerid"));

                            string SchemaName = entity.GetAttributeValue<string>("hil_schemaname").ToString();
                            string RecordType = entity.FormattedValues["hil_recordtype"].ToString();
                            string newvalue = entity.GetAttributeValue<string>("hil_newvalue");

                            EntityReference entityrecord = null;
                            Entity tender = null;
                            if (entity.Contains("hil_tenderno"))
                            {
                                entityrecord = entity.GetAttributeValue<EntityReference>("hil_tenderno");
                                tender = new Entity(entityrecord.LogicalName);
                                tender.Id = entityrecord.Id;
                            }
                            else if (entity.Contains("hil_orderchecklist"))
                            {
                                entityrecord = entity.GetAttributeValue<EntityReference>("hil_orderchecklist");
                                tender = new Entity(entityrecord.LogicalName);
                                tender.Id = entityrecord.Id;
                            }
                            tracingService.Trace("entityrecord.LogicalName " + entityrecord.LogicalName);
                            tracingService.Trace("RecordType " + RecordType);
                            tracingService.Trace("SchemaName " + SchemaName);

                            if (HelperClass.getUserSecurityRole(user, service, "Production Approver", tracingService))
                            {

                                RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
                                {
                                    EntityLogicalName = entityrecord.LogicalName,
                                    LogicalName = SchemaName,
                                    RetrieveAsIfPublished = false
                                };
                                RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);

                                tracingService.Trace("Retrieved the attribute {0}.",
                                    attributeResponse.AttributeMetadata.SchemaName);
                                var attri = attributeResponse.AttributeMetadata;
                                if (attri.AttributeType == AttributeTypeCode.Picklist)
                                {
                                    string optionLable = string.Empty;
                                    var option = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)attri).OptionSet.Options;
                                    int value = -1;
                                    for (int j = 0; j < option.Count; j++)
                                    {
                                        optionLable += option[j].Label.LocalizedLabels[0].Label + " ,";
                                        //throw new InvalidPluginExecutionException("Convert.ToInt32(option[j].Value); " + Convert.ToInt32(option[j].Value));
                                        if (option[j].Label.LocalizedLabels[0].Label.ToString().ToLower() == newvalue.ToLower())
                                        {
                                            value = Convert.ToInt32(option[j].Value);

                                            tender[SchemaName] = new OptionSetValue(value);
                                            //throw new InvalidPluginExecutionException("newvalue " + newvalue + " value " + value);
                                        }
                                    }
                                    if (value == -1)
                                    {
                                        throw new InvalidPluginExecutionException("Please enter correct value out of these " + optionLable);
                                    }
                                }
                                else if (attri.AttributeType == AttributeTypeCode.Lookup)
                                {
                                    var TargetEntity = ((Microsoft.Xrm.Sdk.Metadata.LookupAttributeMetadata)attri).Targets;
                                    //throw new InvalidPluginExecutionException("TargetEntity" + TargetEntity[0].ToString());
                                    //throw new InvalidPluginExecutionException("newvalue" + newvalue);
                                    QueryExpression targetLook = new QueryExpression(TargetEntity[0].ToString());
                                    targetLook.ColumnSet = new ColumnSet(false);
                                    targetLook.Criteria = new FilterExpression(LogicalOperator.And);
                                    targetLook.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, newvalue);
                                    EntityCollection entity1 = service.RetrieveMultiple(targetLook);

                                    if (entity1.Entities.Count > 0)
                                    {
                                        tracingService.Trace("Customerid " + entity1.Entities[0].Id);

                                        tender[SchemaName] = new EntityReference("account", entity1.Entities[0].Id);
                                        tracingService.Trace("CustomerName " + entity1.Entities[0].Id);

                                    }
                                    else
                                    {
                                        throw new InvalidPluginExecutionException("Please enter correct Customer Code  ");
                                    }
                                }
                                else
                                    tender[SchemaName] = newvalue;
                                service.Update(tender);
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException("You do not have permission to make changes. Please contact to Administrator");
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
            #endregion GetField

        }
        void SetApprover(IOrganizationService service, Entity entity, ITracingService tracingService)
        {

            EntityReference targetOwner = entity.GetAttributeValue<EntityReference>("ownerid");

            QueryExpression qry = new QueryExpression("hil_approvalmatrix");
            qry.ColumnSet = new ColumnSet(true);
            qry.Criteria = new FilterExpression(LogicalOperator.And);
            qry.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Change Request Approver");
            qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection ApprovalMatrixColl = service.RetrieveMultiple(qry);
            tracingService.Trace("ApprovalMatrixColl  Count " + ApprovalMatrixColl.Entities.Count);

            QueryExpression _query = new QueryExpression("hil_userbranchmapping");
            _query.ColumnSet = new ColumnSet("hil_name", "hil_zonalhead", "hil_user", "hil_salesoffice", "hil_buhead", "hil_branchproducthead", "hil_scm");
            _query.Criteria = new FilterExpression(LogicalOperator.And);
            _query.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, targetOwner.Id));
            _query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection userMapingColl = service.RetrieveMultiple(_query);

            tracingService.Trace("userMapingColl " + userMapingColl.Entities.Count);
            String approvalPosition = ApprovalMatrixColl[0].GetAttributeValue<EntityReference>("hil_approverposition").Name;
            tracingService.Trace("approvalPosition " + approvalPosition);
            EntityReference approver = new EntityReference();
            if (userMapingColl.Entities.Count > 0)
                approver = HavellsNewPlugin.Approval.ApprovalHelper.getApproverByPosition(userMapingColl[0], approvalPosition, service, tracingService);
            else
                throw new InvalidPluginExecutionException("Approver not Found");

            tracingService.Trace("approver.Id " + approver.Id);

            Entity entity1 = new Entity(entity.LogicalName, entity.Id);
            entity1["hil_crapprover"] = new EntityReference("systemuser", approver.Id);

            service.Update(entity1);


        }

    }
}
