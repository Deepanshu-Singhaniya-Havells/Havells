using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin
{
    public class HelperWOProduct
    {
        public static void SetWarrantyStatus(Entity entity, IOrganizationService service)
        {
            try
            {
                msdyn_workorderincident woIncident = (msdyn_workorderincident)service.Retrieve(entity.LogicalName, entity.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("msdyn_customerasset", "msdyn_workorder"));
                if (woIncident.msdyn_CustomerAsset != null)
                {
                    msdyn_workorder iWo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, woIncident.msdyn_WorkOrder.Id, new ColumnSet("createdon"));
                    msdyn_customerasset enAsset = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, woIncident.msdyn_CustomerAsset.Id,new ColumnSet("hil_warrantytilldate", "hil_warrantystatus", "hil_invoicedate"));
                    msdyn_workorderincident upWOIncident = new msdyn_workorderincident();
                    hil_inventorytype result = hil_inventorytype.NANotfound;
                    OptionSetValue _asWty = new OptionSetValue();
                    if(enAsset.Attributes.Contains("hil_warrantystatus"))
                    {
                        _asWty = (OptionSetValue)enAsset["hil_warrantystatus"];
                    }
                    if(_asWty.Value != 3)
                    {
                        if (enAsset.hil_InvoiceDate != null)
                        {
                            //hil_inventorytype result = HelperWarranty.GetWarrantyStatus(service, woIncident.msdyn_CustomerAsset.Id);
                            DateTime dtEndDate = (DateTime)enAsset["hil_warrantytilldate"];//HelperWarranty.GetWarrantyEndDate(service, woIncident.msdyn_CustomerAsset.Id);
                            DateTime today = (DateTime)iWo.CreatedOn;
                            if (dtEndDate != null)
                            {
                                upWOIncident["new_warrantyenddate"] = dtEndDate;
                                if (dtEndDate.Date >= today.Date)
                                {
                                    result = hil_inventorytype.InWarranty;
                                    upWOIncident.hil_warrantystatus = new OptionSetValue(1);//in

                                }
                                else upWOIncident.hil_warrantystatus = new OptionSetValue(2);//out
                            }
                            else
                            {
                                upWOIncident.hil_warrantystatus = new OptionSetValue(2);//NA not found
                            }
                        }
                        else
                        {
                            upWOIncident.hil_warrantystatus = new OptionSetValue(2);//out
                        }
                    }
                    else
                    {
                        upWOIncident.hil_warrantystatus = new OptionSetValue(2);//Warranty Void
                    }
                    upWOIncident.msdyn_workorderincidentId = woIncident.msdyn_workorderincidentId.Value;
                    service.Update(upWOIncident);
                    #region WorkOrderWarrantyStatusChange
                    if (result == hil_inventorytype.InWarranty)
                    {
                        if (woIncident.msdyn_WorkOrder != null)
                        {
                            msdyn_workorder upWo = new msdyn_workorder();
                            upWo.msdyn_workorderId = woIncident.msdyn_WorkOrder.Id;
                            upWo.hil_WarrantyStatus = new OptionSetValue(1); //In warranty
                            service.Update(upWo);
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelperWOProduct.SetWarrantyStatus" + ex.Message);
            }
        }

        //public static void SetWarrantyStatus(Entity entity, IOrganizationService service)
        //{
        //    try
        //    {
        //        msdyn_workorderincident woIncident = (msdyn_workorderincident)service.Retrieve(entity.LogicalName, entity.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("msdyn_customerasset", "msdyn_workorder"));
        //        if (woIncident.msdyn_CustomerAsset != null)
        //        {
        //            //hil_inventorytype result = HelperWarranty.GetWarrantyStatus(service, woIncident.msdyn_CustomerAsset.Id);
        //            DateTime dtEndDate = HelperWarranty.GetWarrantyEndDate(service, woIncident.msdyn_CustomerAsset.Id);

        //            hil_inventorytype result = hil_inventorytype.NANotfound;
        //            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.ToUniversalTime().Day);


        //            msdyn_workorderincident upWOIncident = new msdyn_workorderincident();
        //            if (dtEndDate != DateTime.MinValue)
        //            {
        //                upWOIncident["new_warrantyenddate"] = dtEndDate;
        //                if (dtEndDate >= today)
        //                {
        //                    result = hil_inventorytype.InWarranty;
        //                }
        //                else
        //                    result = hil_inventorytype.OutWarranty;
        //            }

        //            upWOIncident.msdyn_workorderincidentId = woIncident.msdyn_workorderincidentId.Value;
        //            if (result == hil_inventorytype.InWarranty)
        //            {
        //                upWOIncident.hil_warrantystatus = new OptionSetValue(1);
        //            }
        //            else if (result == hil_inventorytype.OutWarranty)
        //            {
        //                upWOIncident.hil_warrantystatus = new OptionSetValue(2);
        //            }
        //            else if (result == hil_inventorytype.NANotfound)
        //            {
        //                upWOIncident.hil_warrantystatus = new OptionSetValue(4);
        //            }

        //            service.Update(upWOIncident);


        //            #region WorkOrderWarrantyStatusChange
        //            if (result == hil_inventorytype.InWarranty)
        //            {
        //                if (woIncident.msdyn_WorkOrder != null)
        //                {
        //                    msdyn_workorder upWo = new msdyn_workorder();
        //                    upWo.msdyn_workorderId = woIncident.msdyn_WorkOrder.Id;
        //                    upWo.hil_WarrantyStatus = new OptionSetValue(1); //In warranty
        //                    service.Update(upWo);
        //                }
        //            }
        //            #endregion
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidPluginExecutionException("Havells_Plugin.HelperWOProduct.SetWarrantyStatus" + ex.Message);
        //    }
        //}
    }
}
