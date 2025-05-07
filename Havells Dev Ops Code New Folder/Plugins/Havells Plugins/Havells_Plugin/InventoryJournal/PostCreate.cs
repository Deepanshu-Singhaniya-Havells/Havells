using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.InventoryJournal
{
    public class PostCreate : IPlugin
    {
        private static ITracingService tracingService = null;

       public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_inventoryjournal.EntityLogicalName.ToLower() && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    InvJrn_InitFun(entity, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.InventoryJournal.PostCreate.Execute" + ex.Message);
            }
        }
        #region Commented
        //public static void InvJrn_InitFun(Entity entity, IOrganizationService service)
        //{
        //    try
        //    {
        //        hil_inventoryjournal enInvJournal = entity.ToEntity<hil_inventoryjournal>();
        //        enInvJournal = (hil_inventoryjournal)service.Retrieve(enInvJournal.LogicalName, enInvJournal.Id, new ColumnSet(true));
        //        tracingService.Trace("2");
        //        tracingService.Trace("Part" + enInvJournal.hil_Part.Id);
        //        tracingService.Trace("owneracc" + enInvJournal.hil_OwnerAccount);
        //        tracingService.Trace("inv type" + enInvJournal.hil_InventoryType.Value);
        //        tracingService.Trace("Transaction type" + enInvJournal.hil_TransactionOrigin.Value);
        //        tracingService.Trace("WO Prod" + enInvJournal.hil_WorkOrderProduct);
        //        tracingService.Trace("GrnLine" + enInvJournal.hil_GrnLine);
        //        tracingService.Trace("Inv req" + enInvJournal.hil_InventoryRequest);
        //        if (enInvJournal.hil_Part == null) throw new InvalidPluginExecutionException("Part Missing in IJ");
        //        if (enInvJournal.hil_OwnerAccount == null) throw new InvalidPluginExecutionException("Account Missing in IJ");
        //        if (enInvJournal.hil_InventoryType == null) throw new InvalidPluginExecutionException("Inventory Type Missing in IJ");
        //        if (enInvJournal.hil_TransactionOrigin == null) throw new InvalidPluginExecutionException("Transaction Origin Missing in IJ");
        //        if (enInvJournal.hil_Part == null) throw new InvalidPluginExecutionException("Part Missing in IJ");
        //        if (!(enInvJournal.hil_WorkOrderProduct != null || enInvJournal.hil_GrnLine != null || enInvJournal.hil_InventoryRequest != null
        //            || enInvJournal.hil_ProductRequest != null || enInvJournal.hil_ReturnLine != null))
        //            throw new InvalidPluginExecutionException("No transaction lookup found in IJ");
        //        if (enInvJournal.hil_Part != null && enInvJournal.hil_OwnerAccount != null && enInvJournal.hil_InventoryType != null && enInvJournal.hil_TransactionOrigin != null
        //            &&(enInvJournal.hil_WorkOrderProduct != null || enInvJournal.hil_GrnLine != null || enInvJournal.hil_InventoryRequest!=null
        //            || enInvJournal.hil_ProductRequest!=null || enInvJournal.hil_ReturnLine != null))
        //        {
        //            tracingService.Trace("3");
        //            #region Variables
        //            Guid fsPart = enInvJournal.hil_Part.Id;
        //            Guid fsOwnerAccount = enInvJournal.hil_OwnerAccount.Id;
        //            Guid fsOwner = enInvJournal.OwnerId.Id;
        //            OptionSetValue opInventoryType = enInvJournal.hil_InventoryType;
        //            int iAvailableQtyChange = 0;
        //            int iDefectiveQtyChange = 0;
        //            int iDamagedQtyChange = 0;
        //            if(enInvJournal.hil_AvailableQtyChange != null)
        //            {
        //                iAvailableQtyChange = enInvJournal.hil_AvailableQtyChange.Value;
        //            }
        //            if(enInvJournal.hil_DefectiveQtyChange != null)
        //            {
        //                iDefectiveQtyChange = enInvJournal.hil_DefectiveQtyChange.Value;
        //            }
        //            if(enInvJournal.hil_DamagedQtyChange != null)
        //            {
        //                iDamagedQtyChange = enInvJournal.hil_DamagedQtyChange.Value;
        //            }
        //            #endregion
        //            Guid fsInventoryId = Guid.Empty;
        //            hil_inventory result = HelperInventory.FindInventoryCommon(fsPart, fsOwnerAccount, fsOwner, service);
        //            tracingService.Trace("4"+ result.Id);
        //            if (result.Id != Guid.Empty)
        //            {
        //                fsInventoryId = result.Id;
        //                tracingService.Trace("4" + "Calling UpdateInventory");
        //                if (iDefectiveQtyChange > 0 || iDamagedQtyChange > 0)
        //                {
        //                    HelperInventory.UpdateInventory(result, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange, service, opInventoryType);
        //                }
        //                else
        //                {
        //                    HelperInventory.UpdateInventory(result, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange, service, opInventoryType);
        //                }
        //                //HelperInventory.UpdateInventory(result, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange, service);
        //            }
        //            else
        //            {
        //                tracingService.Trace("5" + "Calling CreateInventory");
        //                fsInventoryId = HelperInventory.CreateInventory(fsPart, fsOwnerAccount, fsOwner, opInventoryType, service, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange);
        //            }
        //            //Tag inventory with Inventory journal
        //            hil_inventoryjournal uInvJournal = new hil_inventoryjournal();
        //            uInvJournal.hil_inventoryjournalId = enInvJournal.Id;
        //            uInvJournal.hil_Inventory = new EntityReference(hil_inventory.EntityLogicalName, fsInventoryId);
        //            service.Update(uInvJournal);
        //            tracingService.Trace("6");
        //        }
        //        else
        //        {
        //            throw new InvalidPluginExecutionException("One of the key fields of inventory journal is missing.");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidPluginExecutionException("Havells_Plugin.Inventory.InventoryJournal.PostCreate.Execute" + ex.Message);
        //    }
        //}
        #endregion
        public static void InvJrn_InitFun(Entity entity, IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                iTrace.Trace(" InvJrn_InitFun - 1");
                hil_inventoryjournal enInvJournal = entity.ToEntity<hil_inventoryjournal>();
                iTrace.Trace(" InvJrn_InitFun - 2");
                enInvJournal = (hil_inventoryjournal)service.Retrieve(enInvJournal.LogicalName, enInvJournal.Id, new ColumnSet(true));
                iTrace.Trace(" InvJrn_InitFun - 3");
                OptionSetValue opInventoryType = new OptionSetValue();
                //tracingService.Trace(" InvJrn_InitFun - 4");
                //tracingService.Trace(" InvJrn_InitFun - Part" + enInvJournal.hil_Part.Id);
                //tracingService.Trace(" InvJrn_InitFun - owneracc" + enInvJournal.hil_OwnerAccount);
                //tracingService.Trace(" InvJrn_InitFun - Transaction type" + enInvJournal.hil_TransactionOrigin.Value);
                //tracingService.Trace(" InvJrn_InitFun - WO Prod" + enInvJournal.hil_WorkOrderProduct);
                //tracingService.Trace(" InvJrn_InitFun - GrnLine" + enInvJournal.hil_GrnLine);
                //tracingService.Trace(" InvJrn_InitFun - Inv req" + enInvJournal.hil_InventoryRequest);
                if (enInvJournal.hil_Part == null) throw new InvalidPluginExecutionException("Part Missing in IJ");
                if (enInvJournal.hil_OwnerAccount == null) throw new InvalidPluginExecutionException("Account Missing in IJ");
                //if (enInvJournal.hil_InventoryType == null) throw new InvalidPluginExecutionException("Inventory Type Missing in IJ");
                if (enInvJournal.hil_TransactionOrigin == null) throw new InvalidPluginExecutionException("Transaction Origin Missing in IJ");
                if (enInvJournal.hil_Part == null) throw new InvalidPluginExecutionException("Part Missing in IJ");
                if (!(enInvJournal.hil_WorkOrderProduct != null || enInvJournal.hil_GrnLine != null || enInvJournal.hil_InventoryRequest != null
                    || enInvJournal.hil_ProductRequest != null || enInvJournal.hil_ReturnLine != null))
                    throw new InvalidPluginExecutionException("No transaction lookup found in IJ");

                if (enInvJournal.hil_Part != null && enInvJournal.hil_OwnerAccount != null && enInvJournal.hil_TransactionOrigin != null
                    && (enInvJournal.hil_WorkOrderProduct != null || enInvJournal.hil_GrnLine != null || enInvJournal.hil_InventoryRequest != null
                    || enInvJournal.hil_ProductRequest != null || enInvJournal.hil_ReturnLine != null))
                {

                    tracingService.Trace(" InvJrn_InitFun - 5");
                    #region Variables
                    Guid fsPart = enInvJournal.hil_Part.Id;
                    Guid fsOwnerAccount = enInvJournal.hil_OwnerAccount.Id;
                    Guid fsOwner = enInvJournal.OwnerId.Id;
                    if (enInvJournal.hil_InventoryType != null)
                    {
                        tracingService.Trace(" InvJrn_InitFun - 6");
                        opInventoryType = enInvJournal.hil_InventoryType;
                    }

                    int iAvailableQtyChange = 0;
                    int iDefectiveQtyChange = 0;
                    int iDamagedQtyChange = 0;

                    if (enInvJournal.hil_AvailableQtyChange != null)
                    {
                        tracingService.Trace(" InvJrn_InitFun - 6");
                        iAvailableQtyChange = enInvJournal.hil_AvailableQtyChange.Value;
                    }
                    if (enInvJournal.hil_DefectiveQtyChange != null)
                    {
                        tracingService.Trace(" InvJrn_InitFun - 7");
                        iDefectiveQtyChange = enInvJournal.hil_DefectiveQtyChange.Value;
                    }
                    if (enInvJournal.hil_DamagedQtyChange != null)
                    {
                        tracingService.Trace(" InvJrn_InitFun - 8");
                        iDamagedQtyChange = enInvJournal.hil_DamagedQtyChange.Value;
                    }
                    #endregion
                    Guid fsInventoryId = Guid.Empty;
                    hil_inventory result = HelperInventory.FindInventoryCommon(fsPart, fsOwnerAccount, fsOwner, service);
                    tracingService.Trace("4" + result.Id);
                    if (result.Id != Guid.Empty)
                    {
                        fsInventoryId = result.Id;
                        tracingService.Trace("4" + "Calling UpdateInventory");
                        if (iDefectiveQtyChange > 0 || iDamagedQtyChange > 0)
                        {
                            tracingService.Trace("4" + "Calling UpdateInventoryDefective");
                            HelperInventory.UpdateInventory(result, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange, service, opInventoryType, iTrace);
                        }
                        else
                        {
                            tracingService.Trace("4" + "Calling UpdateInventoryGood");
                            HelperInventory.UpdateInventory(result, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange, service, opInventoryType, iTrace);
                        }
                    }
                    else
                    {
                        tracingService.Trace("5" + "Calling CreateInventory");

                        fsInventoryId = HelperInventory.CreateInventory(fsPart, fsOwnerAccount, fsOwner, opInventoryType, service, iAvailableQtyChange, iDefectiveQtyChange, iDamagedQtyChange);
                    }
                    //Tag inventory with Inventory journal
                    hil_inventoryjournal uInvJournal = new hil_inventoryjournal();
                    uInvJournal.hil_inventoryjournalId = enInvJournal.Id;
                    uInvJournal.hil_Inventory = new EntityReference(hil_inventory.EntityLogicalName, fsInventoryId);
                    service.Update(uInvJournal);
                    tracingService.Trace("6");
                }
                else
                {
                    throw new InvalidPluginExecutionException("One of the key fields of inventory journal is missing.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Inventory.InventoryJournal.PostCreate.Execute" + ex.Message);
            }
        }
    }
}