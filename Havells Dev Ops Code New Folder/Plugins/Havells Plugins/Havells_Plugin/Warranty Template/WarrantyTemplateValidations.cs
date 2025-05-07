using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using System.Linq;

namespace Havells_Plugin.WarrantyTemplate
{
    public class WarrantyTemplateValidations : IPlugin
    {
        private IOrganizationService service;
        private ITracingService tracingService = null;
        private Guid ProductSubCategory = Guid.Empty;
        private int Type = -1;
        private Guid Plan = Guid.Empty;
        private DateTime f = new DateTime(1900, 1, 1);
        private DateTime t = new DateTime(1900, 1, 1);
        private string updateCondition = string.Empty;

        private bool IsDuplicate(DateTime f, DateTime t, DateTime f1, DateTime t1)
        {
            bool toReturn = false;
            if (f1 >= f && f1 <= t)
            {
                toReturn = true;
            }
            else if (t1 >= f && t1 <= t)
            {
                toReturn = true;
            }
            else if (f1 <= f && t1 >= t)
            {
                toReturn = true;
            }
            return toReturn;
        }

        private List<Entity> FetchSimilar()
        {
            
            string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_warrantytemplate'>
                <attribute name='hil_name' />
                <attribute name='hil_validto' />
                <attribute name='hil_validfrom' />
                <filter type='and'>
                <condition attribute='hil_product' operator='eq' value='{ProductSubCategory}'/>
                <condition attribute='hil_type' operator='eq' value='{Type}'/>
                <condition attribute='statecode' operator='eq' value='0' />"
                + updateCondition +
                @"</filter>
                </entity>
                </fetch>";

            return service.RetrieveMultiple(new FetchExpression(fetch)).Entities.ToList();
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                #region PluginConfig
                tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = serviceFactory.CreateOrganizationService(context.UserId);
                #endregion

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_warrantytemplate" && context.Depth == 1)
                {
                    Entity warrantyTemplate = (Entity)context.InputParameters["Target"];
                    if (context.MessageName.ToUpper() == "UPDATE")
                    {
                        warrantyTemplate = service.Retrieve("hil_warrantytemplate", warrantyTemplate.Id, new ColumnSet(true));
                        updateCondition = $"<condition attribute='hil_warrantytemplateid' operator='ne' value='{warrantyTemplate.Id}' />";
                    }
                    if (warrantyTemplate.Contains("hil_product")) ProductSubCategory = warrantyTemplate.GetAttributeValue<EntityReference>("hil_product").Id;
                    if (warrantyTemplate.Contains("hil_type")) Type = warrantyTemplate.GetAttributeValue<OptionSetValue>("hil_type").Value;
                    if (warrantyTemplate.Contains("hil_amcplan"))
                    {
                        Plan = warrantyTemplate.GetAttributeValue<EntityReference>("hil_amcplan").Id;
                        updateCondition += $"<condition attribute='hil_amcplan' operator='eq' value='{Plan}'/>";
                    }
                    if (warrantyTemplate.Contains("hil_validfrom")) f = warrantyTemplate.GetAttributeValue<DateTime>("hil_validfrom").Date;
                    if (warrantyTemplate.Contains("hil_validto")) t = warrantyTemplate.GetAttributeValue<DateTime>("hil_validto").Date;

                    List<Entity> data = FetchSimilar();
                    string _warrantyTemplates = string.Empty;
                    for (int i = 0; i < data.Count; i++)
                    {
                        DateTime f1 = data[i].GetAttributeValue<DateTime>("hil_validfrom").Date;
                        DateTime t1 = data[i].GetAttributeValue<DateTime>("hil_validto").Date;
                        string templateId = data[i].GetAttributeValue<string>("hil_name");
                        if (IsDuplicate(f, t, f1, t1))
                            _warrantyTemplates += templateId + ",";
                    }
                    if (!string.IsNullOrWhiteSpace(_warrantyTemplates))
                    {
                        throw new InvalidPluginExecutionException("Overlapping warranty template found with template Id: " + _warrantyTemplates);
                    }

                    //Validate Warranty Template Valid From and Valid Until 
                    ValidateDateRange(service, context);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("ERROR!!! " + ex.Message);
            }
        }

        public void ValidateDateRange(IOrganizationService service, IPluginExecutionContext context) {
            DateTime validFrom = new DateTime(1990, 1, 2), validTo = new DateTime(1990, 1, 1);
            Entity entity = (Entity)context.InputParameters["Target"];
            if (context.MessageName == "Update")
            {
                Entity e2 = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_validfrom", "hil_validto"));
                if (entity.Contains("hil_validfrom"))
                {
                    validFrom = entity.GetAttributeValue<DateTime>("hil_validfrom");
                    validTo = e2.GetAttributeValue<DateTime>("hil_validto");
                }

                if (entity.Contains("hil_validto"))
                {
                    validTo = entity.GetAttributeValue<DateTime>("hil_validto");
                    validFrom = e2.GetAttributeValue<DateTime>("hil_validfrom");
                }

            }
            if (context.MessageName == "Create")
            {
                if (entity.Contains("hil_validfrom") && entity.Contains("hil_validto"))
                {
                    validFrom = entity.GetAttributeValue<DateTime>("hil_validfrom");
                    validTo = entity.GetAttributeValue<DateTime>("hil_validto");
                }
            }
            if (validFrom > validTo)
            {
                throw new InvalidPluginExecutionException("Start Date of validity should be less than or equal to the ending date of validity");
            }
        }
    }
}
