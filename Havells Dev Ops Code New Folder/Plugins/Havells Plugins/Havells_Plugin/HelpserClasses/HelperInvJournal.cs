using System;
using Microsoft.Xrm.Sdk;

namespace Havells_Plugin
{
    public class HelperInvJournal
    {
        public static Guid CreateInventoryJournal(Guid fsPart, Guid fsOwnerAccount, Guid fsOwner, Guid fsWOProductId, Guid fsGRNLineId, Guid fsInvReqId, Guid fsProductRequest, Guid fsReturnLineId, IOrganizationService service, ITracingService tracing, Int32 iAvailableQtyChange = 0, Int32 iDefectiveQtyChange = 0)
        {
            try
            {
                hil_inventoryjournal crInvJournal = new hil_inventoryjournal();
                crInvJournal.hil_Part = new EntityReference(Product.EntityLogicalName, fsPart);
                crInvJournal.hil_OwnerAccount = new EntityReference(Account.EntityLogicalName, fsOwnerAccount);
                crInvJournal.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsOwner);
                //crInvJournal.hil_InventoryType = opInventoryType;
                crInvJournal.hil_AvailableQtyChange = iAvailableQtyChange;
                //crInvJournal.hil_DamagedQtyChange = iDamagedQtyChange;
                //crInvJournal.hil_DefectiveQtyChange = iDefectiveQtyChange;
                if (fsWOProductId != Guid.Empty)
                {
                    crInvJournal.hil_WorkOrderProduct = new EntityReference(msdyn_workorderproduct.EntityLogicalName, fsWOProductId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(1);//Work Order
                }
                else if (fsGRNLineId != Guid.Empty)
                {
                    crInvJournal.hil_GrnLine = new EntityReference(hil_grnline.EntityLogicalName, fsGRNLineId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(2);//GRN
                }
                else if (fsInvReqId != Guid.Empty)
                {
                    crInvJournal.hil_InventoryRequest = new EntityReference(hil_inventoryrequest.EntityLogicalName, fsInvReqId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(3);//Inv request
                }
                else if (fsProductRequest != Guid.Empty)
                {
                    crInvJournal.hil_ProductRequest = new EntityReference(hil_productrequest.EntityLogicalName, fsProductRequest);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(4);//Product request
                }
                else if (fsReturnLineId != Guid.Empty)
                {
                    crInvJournal.hil_ReturnLine = new EntityReference(hil_ReturnLine.EntityLogicalName, fsReturnLineId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(5);//hil_ReturnLine
                }
                Guid journalId = service.Create(crInvJournal);
                return journalId;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperInvJournal.CreateInvJournal" + ex.Message);
            }
        }
        public static void CreateInvJournal(Guid fsPart, Guid fsOwnerAccount, Guid fsOwner, Guid fsWOProductId, Guid fsGRNLineId, Guid fsInvReqId, Guid fsProductRequest, Guid fsReturnLineId, IOrganizationService service, ITracingService tracing, Int32 iAvailableQtyChange = 0, Int32 iDefectiveQtyChange = 0)
        {
            try
            {
                hil_inventoryjournal crInvJournal = new hil_inventoryjournal();
                crInvJournal.hil_Part = new EntityReference(Product.EntityLogicalName, fsPart);
                crInvJournal.hil_OwnerAccount = new EntityReference(Account.EntityLogicalName, fsOwnerAccount);
                crInvJournal.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsOwner);
                //crInvJournal.hil_InventoryType = opInventoryType;
                crInvJournal.hil_AvailableQtyChange = iAvailableQtyChange;
                //crInvJournal.hil_DamagedQtyChange = iDamagedQtyChange;
                //crInvJournal.hil_DefectiveQtyChange = iDefectiveQtyChange;
                if (fsWOProductId != Guid.Empty)
                {
                    crInvJournal.hil_WorkOrderProduct = new EntityReference(msdyn_workorderproduct.EntityLogicalName, fsWOProductId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(1);//Work Order
                }
                else if (fsGRNLineId != Guid.Empty)
                {
                    crInvJournal.hil_GrnLine = new EntityReference(hil_grnline.EntityLogicalName, fsGRNLineId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(2);//GRN
                }
                else if (fsInvReqId != Guid.Empty)
                {
                    crInvJournal.hil_InventoryRequest = new EntityReference(hil_inventoryrequest.EntityLogicalName, fsInvReqId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(3);//Inv request
                }
                else if (fsProductRequest != Guid.Empty)
                {
                    crInvJournal.hil_ProductRequest = new EntityReference(hil_productrequest.EntityLogicalName, fsProductRequest);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(4);//Product request
                }
                else if (fsReturnLineId != Guid.Empty)
                {
                    crInvJournal.hil_ReturnLine = new EntityReference(hil_ReturnLine.EntityLogicalName, fsReturnLineId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(5);//hil_ReturnLine
                }
                service.Create(crInvJournal);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperInvJournal.CreateInvJournal" + ex.Message);
            }
        }
        public static void CreateInvJournalDefective(Guid fsPart, Guid fsOwnerAccount, Guid fsOwner, OptionSetValue opInventoryType, Guid fsWOProductId, Guid fsGRNLineId, Guid fsInvReqId, Guid fsProductRequest, Guid fsReturnLineId, IOrganizationService service, ITracingService iTrace, Int32 iDefectiveQtyChange = 0)
        {
            try
            {
                iTrace.Trace("CreateInvJournalDefective - 1");
                hil_inventoryjournal crInvJournal = new hil_inventoryjournal();
                crInvJournal.hil_Part = new EntityReference(Product.EntityLogicalName, fsPart);
                crInvJournal.hil_OwnerAccount = new EntityReference(Account.EntityLogicalName, fsOwnerAccount);
                crInvJournal.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsOwner);
                crInvJournal.hil_InventoryType = opInventoryType;
                //crInvJournal.hil_AvailableQtyChange = iAvailableQtyChange;
                //crInvJournal.hil_DamagedQtyChange = iDamagedQtyChange;
                //if(opInventoryType != null && opInventoryType.Value == 1)
                //{
                crInvJournal.hil_DefectiveQtyChange = iDefectiveQtyChange;
                //}
                //else
                //{
                //    crInvJournal["hil_outwarrantydefectivequantity"] = Convert.ToInt32(iDefectiveQtyChange);
                //}

                if (fsWOProductId != Guid.Empty)
                {
                    iTrace.Trace("CreateInvJournalDefective - 2");
                    crInvJournal.hil_WorkOrderProduct = new EntityReference(msdyn_workorderproduct.EntityLogicalName, fsWOProductId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(1);//Work Order
                }
                else if (fsGRNLineId != Guid.Empty)
                {
                    iTrace.Trace("CreateInvJournalDefective - 3");
                    crInvJournal.hil_GrnLine = new EntityReference(hil_grnline.EntityLogicalName, fsGRNLineId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(2);//GRN
                }
                else if (fsInvReqId != Guid.Empty)
                {
                    iTrace.Trace("CreateInvJournalDefective - 4");
                    crInvJournal.hil_InventoryRequest = new EntityReference(hil_inventoryrequest.EntityLogicalName, fsInvReqId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(3);//Inv request
                }
                else if (fsProductRequest != Guid.Empty)
                {
                    iTrace.Trace("CreateInvJournalDefective - 5");
                    crInvJournal.hil_ProductRequest = new EntityReference(hil_productrequest.EntityLogicalName, fsProductRequest);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(4);//Product request
                }
                else if (fsReturnLineId != Guid.Empty)
                {
                    iTrace.Trace("CreateInvJournalDefective - 6");
                    crInvJournal.hil_ReturnLine = new EntityReference(hil_ReturnLine.EntityLogicalName, fsReturnLineId);
                    crInvJournal.hil_TransactionOrigin = new OptionSetValue(5);//hil_ReturnLine
                }
                service.Create(crInvJournal);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperInvJournal.CreateInvJournal" + ex.Message);
            }
        }
    }
}
