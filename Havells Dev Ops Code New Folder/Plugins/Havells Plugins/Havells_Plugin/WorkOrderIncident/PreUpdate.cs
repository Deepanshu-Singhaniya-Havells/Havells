using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace Havells_Plugin.WorkOrderIncident
{
    public class PreUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                msdyn_workorderincident _entWOInc = entity.ToEntity<msdyn_workorderincident>();

                Entity preImage = (Entity)context.PreEntityImages["image"];
                msdyn_workorderincident PreImageIncident = preImage.ToEntity<msdyn_workorderincident>();

                if (_entWOInc.hil_productreplacement != null && _entWOInc.hil_productreplacement.Value == 1)
                {
                    msdyn_workorder iJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, PreImageIncident.msdyn_WorkOrder.Id, new ColumnSet("hil_incidentquantity", "hil_callsubtype"));

                    if (iJob.hil_CallSubType != null && iJob.hil_CallSubType.Name == "Dealer Stock Repair")
                    {
                        throw new InvalidPluginExecutionException("Product Replacement is not allowed on Dealer Stock Call");
                    }
                }

                if (_entWOInc.Contains("msdyn_customerasset"))
                {
                    msdyn_customerasset cus_Assest = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, _entWOInc.msdyn_CustomerAsset.Id, new ColumnSet(new string[] { "msdyn_product", "hil_productcategory", "hil_invoicedate", "hil_invoiceno", "hil_productsubcategorymapping" }));

                    if (cus_Assest.hil_ProductCategory == null || cus_Assest.msdyn_Product == null || cus_Assest.hil_productsubcategorymapping == null
                        || cus_Assest.hil_ProductCategory.Id != PreImageIncident.hil_ProductCategory.Id)
                    {
                        throw new InvalidPluginExecutionException("The Customer Asset category combination should match with Job Incident.");
                    }
                    hil_stagingdivisonmaterialgroupmapping sdmMapping = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(
                            hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, cus_Assest.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Id, new ColumnSet(new string[] { "hil_productsubcategorymg" }));

                    if (sdmMapping.hil_ProductSubCategoryMG == null
                        || sdmMapping.hil_ProductSubCategoryMG.Id != PreImageIncident.hil_ProductSubCategory.Id)
                    {
                        throw new InvalidPluginExecutionException("Product Sub Category should match with assest Product Sub Category. Please contact to Administrator");
                    }
                }
                if (_entWOInc.Contains("hil_quantity"))
                {
                    Product prod_Category = (Product)service.Retrieve(Product.EntityLogicalName, PreImageIncident.hil_ProductSubCategory.Id, new ColumnSet(new string[] { "hil_isserialized" }));
                    if (prod_Category.hil_IsSerialized != null && prod_Category.hil_IsSerialized.Value == 1
                        && (_entWOInc.hil_Quantity == null || _entWOInc.hil_Quantity != 1))
                    {
                        throw new InvalidPluginExecutionException("Quantity should be equals to 1 for serialized Product");
                    }
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PreUpdate" + e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PreUpdate" + e.Message);
            }
        }
    }
}
