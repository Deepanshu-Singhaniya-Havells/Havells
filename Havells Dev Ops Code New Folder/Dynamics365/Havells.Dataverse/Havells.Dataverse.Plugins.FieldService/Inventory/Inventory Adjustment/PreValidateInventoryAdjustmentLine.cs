using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidateInventoryAdjustmentLine : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion


            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_inventoryspareadjustmentline" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    
                    if (!entity.Contains("hil_adjustmentnumber"))
                        throw new InvalidPluginExecutionException("Adjustment Number is required.");

                    Entity entityHeader = service.Retrieve("hil_inventoryspareadjustment", entity.GetAttributeValue<EntityReference>("hil_adjustmentnumber").Id, new ColumnSet("hil_adjustmenttype", "hil_brand"));
                    OptionSetValue _adjustmentType = entityHeader.GetAttributeValue<OptionSetValue>("hil_adjustmenttype");
                    OptionSetValue _partBrand = entityHeader.GetAttributeValue<OptionSetValue>("hil_brand");

                    if (_adjustmentType.Value == 1)//Stock Opening
                    {
                        if (!entity.Contains("hil_partcode"))
                            throw new InvalidPluginExecutionException("Spare Part is required.");
                        if (!entity.Contains("hil_ajustmentquantity"))
                            throw new InvalidPluginExecutionException("Adjustment Quantity is required.");
                    }
                    else if (_adjustmentType.Value == 2)//Stock Audit
                    {
                    }
    
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='product'>
                    <attribute name='productid' />
                    <order attribute='productnumber' descending='false' />
                    <filter type='and'>
                        <condition attribute='productid' operator='eq' value='{entity.GetAttributeValue<EntityReference>("hil_partcode").Id}' />
                    </filter>
                    <link-entity name='product' from='productid' to='hil_division' visible='false' link-type='outer' alias='div'>
                        <attribute name='hil_brandidentifier' />
                    </link-entity>
                    </entity>
                    </fetch>";

                    EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));

                    if (_entCol.Entities.Count > 0)
                    {
                        OptionSetValue _diviionBrand = (OptionSetValue)_entCol.Entities[0].GetAttributeValue<AliasedValue>("div.hil_brandidentifier").Value;
                        if (_partBrand.Value != _diviionBrand.Value)
                            throw new InvalidPluginExecutionException("Brand Validation!!!\n Part Code's Brand doesn't match with Inventory Adjustment Brand.\n ");
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Brand is not defined in Part Code Division.");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
