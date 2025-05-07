using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin
{
    public class HelperInventory
    {
        public static hil_inventory FindInventory(Guid fsPart, Guid fsOwnerAccount, Guid fsOwner, OptionSetValue opInventoryType, IOrganizationService service)
        {
            hil_inventory result = new hil_inventory();
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _inv in orgContext.CreateQuery<hil_inventory>()
                              where _inv.hil_Part.Id == fsPart
                              && _inv.hil_OwnerAccount.Id == fsOwnerAccount
                              && _inv.OwnerId.Id == fsOwner
                              && _inv.hil_InventoryType == opInventoryType
                              select new
                              {
                                  _inv.hil_inventoryId,
                                  _inv.hil_AvailableQty,
                                  _inv.hil_DamagedQty,
                                  _inv.hil_DefectiveQty,
                                  _inv.hil_name
                              };
                    foreach (var iobj in obj)
                    {
                        result.hil_inventoryId = iobj.hil_inventoryId;
                        result.hil_AvailableQty = iobj.hil_AvailableQty;
                        result.hil_DamagedQty = iobj.hil_DamagedQty;
                        result.hil_DefectiveQty = iobj.hil_DefectiveQty;
                        result.hil_name = iobj.hil_name;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperInventory.FindInventory" + ex.Message);
            }
            return result;
        }
        public static void UpdateInventory(hil_inventory Inventory, Int32 iAvailableQtyChange, Int32 iDefectiveQtyChange, Int32 iDamagedQtyChange, IOrganizationService service, OptionSetValue iOp, ITracingService iTrace)
        {
            String Message = "";
            Int32 IOutWarrantyDef = 0;
            try
            {
                iTrace.Trace("UpdateInventory - 1");
                Message = "Inventory Name" + Inventory.hil_name + " AvailableQtyChange" + iAvailableQtyChange + " DefectiveQtyChange" + iDefectiveQtyChange
                   + "Current Inv Available" + Inventory.hil_AvailableQty.Value;
                if (Inventory.hil_AvailableQty != null) iAvailableQtyChange += Inventory.hil_AvailableQty.Value;
                //if (Inventory.hil_DefectiveQty != null) iDefectiveQtyChange += Inventory.hil_DefectiveQty.Value;
                if (iAvailableQtyChange < 0) throw new InvalidPluginExecutionException("------>>>Inventory Good quantity can't be reduced below 0. Please check inventory quantity for this part.<<<------");
                if (iDefectiveQtyChange < 0) throw new InvalidPluginExecutionException("------>>>Inventory Defective quantity can't be reduced below 0. Please check inventory quantity for this part.<<<------");
                Inventory.hil_AvailableQty = iAvailableQtyChange;
                //Inventory.hil_DamagedQty = Inventory.hil_DamagedQty + iDamagedQtyChange;
                if (iDefectiveQtyChange > 0)
                {
                    iTrace.Trace("UpdateInventory - 2");
                    if (iOp.Value == 1)
                    {
                        iTrace.Trace("UpdateInventory - 3");
                        if(Inventory.hil_DefectiveQty != null)
                        {
                            iDefectiveQtyChange += Inventory.hil_DefectiveQty.Value;
                            Inventory.hil_DefectiveQty = iDefectiveQtyChange;
                        }
                        else
                        {
                            Inventory.hil_DefectiveQty = iDefectiveQtyChange;
                        }
                    }
                    else
                    {
                        iTrace.Trace("UpdateInventory - 4");
                        if (Inventory.Contains("hil_outwarrantydefectivequantity") && Inventory.Attributes.Contains("hil_outwarrantydefectivequantity"))
                        {
                            iTrace.Trace("UpdateInventory - 5");
                            IOutWarrantyDef = Inventory.GetAttributeValue<Int32>("hil_outwarrantydefectivequantity");
                            if(IOutWarrantyDef != null)
                            {
                                iDefectiveQtyChange = IOutWarrantyDef + iDamagedQtyChange;
                                Inventory["hil_outwarrantydefectivequantity"] = iDefectiveQtyChange;
                            }
                            else
                            {
                                Inventory["hil_outwarrantydefectivequantity"] = iDefectiveQtyChange;
                            }
                        }   
                        else
                        {
                            iTrace.Trace("UpdateInventory - 6");
                            Inventory["hil_outwarrantydefectivequantity"] = iDefectiveQtyChange;
                        }
                    }
                }
                service.Update(Inventory);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperInventory.UpdateInventry" + ex.Message + "Message" + Message);
            }
        }
        public static Guid CreateInventory(Guid fsPart, Guid fsOwnerAccount, Guid fsOwner, OptionSetValue opInventoryType, IOrganizationService service, Int32 iAvailableQtyChange = 0, Int32 iDefectiveQtyChange = 0, Int32 iDamagedQtyChange = 0)
        {
            try
            {
                hil_inventory cInventory = new hil_inventory();
                cInventory.hil_Part = new EntityReference(Product.EntityLogicalName, fsPart);
                cInventory.hil_OwnerAccount = new EntityReference(Account.EntityLogicalName, fsOwnerAccount);
                cInventory.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsOwner);
                //cInventory.hil_InventoryType = opInventoryType;
                cInventory.hil_AvailableQty = iAvailableQtyChange;
                cInventory.hil_DamagedQty = iDamagedQtyChange;
                if (iDefectiveQtyChange > 0)
                {
                    if (opInventoryType.Value == 1)
                    {
                        cInventory.hil_DefectiveQty = iDefectiveQtyChange;
                    }
                    else
                    {
                        cInventory["hil_outwarrantydefectivequantity"] = iDefectiveQtyChange;
                    }
                }
                Guid fsInventoryId = service.Create(cInventory);
                return fsInventoryId;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("avells_Plugin.HelpserClasses.HelperInventory.UpdateInventry" + ex.Message);
            }
        }
        public static hil_inventory FindInventoryCommon(Guid fsPart, Guid fsOwnerAccount, Guid fsOwner, IOrganizationService service)
        {
            hil_inventory result = new hil_inventory();
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _inv in orgContext.CreateQuery<hil_inventory>()
                              where _inv.hil_Part.Id == fsPart
                              && _inv.hil_OwnerAccount.Id == fsOwnerAccount
                              && _inv.OwnerId.Id == fsOwner
                              orderby _inv.CreatedOn descending
                              select new
                              {
                                  _inv.hil_inventoryId,
                                  _inv.hil_AvailableQty,
                                  _inv.hil_DamagedQty,
                                  _inv.hil_DefectiveQty,
                                  _inv.hil_name
                              };
                    foreach (var iobj in obj)
                    {
                        result.hil_inventoryId = iobj.hil_inventoryId;
                        result.hil_AvailableQty = iobj.hil_AvailableQty;
                        result.hil_DamagedQty = iobj.hil_DamagedQty;
                        result.hil_DefectiveQty = iobj.hil_DefectiveQty;
                        result.hil_name = iobj.hil_name;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperInventory.FindInventory" + ex.Message);
            }
            return result;
        }
    }
}
