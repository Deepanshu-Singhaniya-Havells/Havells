using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.ClaimOverHeadline
{
   public class ClaimPostingSetupMessage : IPlugin
    {
        public static ITracingService tracingService = null;
        Guid ProductcategoryID = Guid.Empty;
        Guid CallsubtypeID = Guid.Empty;

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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    tracingService.Trace("Checking existence of productcategory and callsubtype for ClaimPostingSetup Start");

                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_productcategory", "hil_callsubtype"));

                    if (entity.Contains("hil_productcategory") && entity.Contains("hil_callsubtype"))
                    {
                        ProductcategoryID = entity.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                        CallsubtypeID = entity.GetAttributeValue<EntityReference>("hil_callsubtype").Id;

                        string _claimpostingsetupFetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_claimpostingsetup'>
                            <attribute name='hil_name' />
                            <attribute name='hil_productcategory' />
                            <attribute name='modifiedon' />
                            <attribute name='modifiedby' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='hil_activitygstslab' />
                            <attribute name='hil_activitycode' />
                            <attribute name='hil_claimpostingsetupid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_productcategory' operator='eq' value='{ProductcategoryID }' />
                              <condition attribute='hil_callsubtype' operator='eq' value='{CallsubtypeID }' />
                            </filter>
                          </entity>
                        </fetch>";

                        EntityCollection _claimpostingsetup = service.RetrieveMultiple(new FetchExpression(_claimpostingsetupFetchXML));
                        tracingService.Trace("Checking existence of productcategory and callsubtype for ClaimPostingSetup End");

                        if (_claimpostingsetup.Entities.Count == 0)
                        {
                            throw new InvalidPluginExecutionException($"Please cross check & correct the Product Category {entity.GetAttributeValue<EntityReference>("hil_productcategory").Name} Or Call Sub Type {entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name}, this combination is not eligible for payment.");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException($"Please cross check & correct the Product Category {entity.GetAttributeValue<EntityReference>("hil_productcategory").Name} Or Call Sub Type {entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name}, this combination is not eligible for payment.");
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }

        }
    }
}
