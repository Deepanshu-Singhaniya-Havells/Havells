using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WarrantyTemplate
{
    public class Warranty_Template_Lines : IPlugin
    {
        private IOrganizationService service;
        private ITracingService tracingService;
        private bool isvalid(EntityReference templateRef, DateTime From, DateTime To)
        {
            Entity warrantyTemplate = service.Retrieve(templateRef.LogicalName, templateRef.Id, new ColumnSet("hil_validfrom", "hil_validto"));
            DateTime validFrom = warrantyTemplate.GetAttributeValue<DateTime>("hil_validfrom");
            DateTime validTo = warrantyTemplate.GetAttributeValue<DateTime>("hil_validto");

            if (From >= validFrom && From <= validTo)
            {
                return false;
            }
            else if (To >= validFrom && To <= validTo) return false;
            else if (From <= validFrom && To >= validTo) return false;
            return true;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            try
            {
                bool IsMessage = context.MessageName == "Create";
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_warrantytemplate.EntityLogicalName && IsMessage)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string serialNumber = string.Empty;
                    EntityReference Model = null;
                    if (entity.Contains("hil_name"))
                    {
                        serialNumber = entity.GetAttributeValue<string>("hil_name");
                    }
                    if (entity.Contains("hil_model"))
                    {
                        Model = entity.GetAttributeValue<EntityReference>("hil_model");
                    }
                    EntityReference templateRef = entity.GetAttributeValue<EntityReference>("hil_warrantytemplate");

                    Entity warrantyTemplate = service.Retrieve(templateRef.LogicalName, templateRef.Id, new ColumnSet("hil_validfrom", "hil_validto"));
                    DateTime validFrom = warrantyTemplate.GetAttributeValue<DateTime>("hil_validfrom");
                    DateTime validTo = warrantyTemplate.GetAttributeValue<DateTime>("hil_validto");

                    QueryExpression query = new QueryExpression("hil_warrantytemplateline");
                    query.ColumnSet = new ColumnSet("hil_warrantytemplate");
                    if (Model != null)
                        query.Criteria.AddCondition("hil_model", ConditionOperator.Equal, Model);
                    if (string.IsNullOrEmpty(serialNumber))
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, serialNumber);
                    EntityCollection tempColl = service.RetrieveMultiple(query);
                    for (int i = 0; i < tempColl.Entities.Count; i++)
                    {
                        EntityReference tempRef = tempColl.Entities[i].GetAttributeValue<EntityReference>("hil_warrantytemplate");
                        if (isvalid(tempRef, validFrom, validTo))
                        {
                            throw new InvalidPluginExecutionException("Overlappping Warranty Template Lines for same Model/Serial Number");
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}