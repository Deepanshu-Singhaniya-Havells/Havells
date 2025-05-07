using Havells.FieldService.RMA.Helper;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Havells.FieldService.RMA.InventoryAdjustment.InventoryAdjustment_Helper;

namespace Havells.FieldService.RMA.Actions
{
    public class PostPurchaseOrderRecipt : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target"))
                {
                    var inParam1 =(EntityReference) context.InputParameters["Target"];
                    Post_PurchaseOrderRecipt(service, inParam1);
                    context.OutputParameters["Remarks"] = "Done";
                    context.OutputParameters["Status"] = true;
                }
                else
                {
                    context.OutputParameters["Remarks"] = "Target Not Found";
                    context.OutputParameters["Status"] = true;
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Remarks"] = ex.Message;
                context.OutputParameters["Status"] = false;
                throw new InvalidPluginExecutionException("Error !!\n " + ex.Message);
            }

        }
        public static void Post_PurchaseOrderRecipt(IOrganizationService service, EntityReference _poReciptEntityRef)
        {
            try
            {
                string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""msdyn_purchaseorderreceiptproduct"">
                                    <attribute name=""hil_billedquantity"" />
                                    <attribute name=""msdyn_name"" />
                                    <attribute name=""hil_sortquantity"" />
                                    <attribute name=""hil_freshquantity"" />
                                    <attribute name=""hil_damagequantity"" />
                                    <attribute name=""msdyn_purchaseorder"" />
                                    <attribute name=""ownerid"" />
                                    <attribute name=""msdyn_purchaseorderreceipt"" />
                                    <attribute name=""msdyn_purchaseorderproduct"" />
                                    <attribute name=""msdyn_associatetoworkorder"" />
                                    <attribute name=""msdyn_associatetowarehouse"" />
                                    <attribute name=""msdyn_purchaseorderreceiptproductid"" />
                                    <filter type=""and"">
                                      <condition attribute=""statuscode"" operator=""eq"" value=""1"" />
                                      <condition attribute=""hil_billedquantity"" operator=""not-null"" />
                                      <condition attribute=""msdyn_purchaseorderreceipt"" operator=""eq"" value=""{_poReciptEntityRef.Id}"" />
                                    </filter>
                                  </entity>
                                </fetch>";
                EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetch));
                int sortQty = 0, freshQty = 0, totalQty = 0, damageQty = 0;
                foreach (Entity entity in collection.Entities)
                {
                    sortQty = 0; freshQty = 0; totalQty = 0; damageQty = 0;
                    freshQty = entity.GetAttributeValue<int>("hil_freshquantity");
                    sortQty = entity.GetAttributeValue<int>("hil_sortquantity");
                    totalQty = entity.GetAttributeValue<int>("hil_billedquantity");
                    damageQty = entity.GetAttributeValue<int>("hil_damagequantity");
                    if ((damageQty + sortQty + freshQty) != totalQty)
                    {
                        throw new InvalidPluginExecutionException("Quantity is mismached for Purchase Order Recipt Product " + entity["msdyn_name"]);
                    }
                }
                foreach (Entity entity in collection.Entities)
                {
                    try
                    {
                        if (entity.Contains("hil_freshquantity"))
                        {
                            Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                            entity1["msdyn_quantity"] = (double)entity.GetAttributeValue<int>("hil_freshquantity");
                            service.Update(entity1);
                        }
                        if (entity.Contains("hil_damagequantity"))
                        {
                            Entity entity1 = new Entity(entity.LogicalName);
                            entity1["msdyn_purchaseorder"] = entity["msdyn_purchaseorder"];
                            entity1["ownerid"] = entity["ownerid"];
                            entity1["msdyn_purchaseorderreceipt"] = entity["msdyn_purchaseorderreceipt"];
                            entity1["msdyn_purchaseorderproduct"] = entity["msdyn_purchaseorderproduct"];
                            entity1["msdyn_associatetoworkorder"] = entity["msdyn_associatetoworkorder"];
                            entity1["msdyn_associatetowarehouse"] = InventoryHelper.RetriveBookableResource(service, entity.GetAttributeValue<EntityReference>("ownerid")).DefectiveWarehouse;
                            entity1["msdyn_quantity"] = (double)entity.GetAttributeValue<int>("hil_damagequantity");
                            service.Create(entity1);
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidPluginExecutionException(ex.Message);
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
