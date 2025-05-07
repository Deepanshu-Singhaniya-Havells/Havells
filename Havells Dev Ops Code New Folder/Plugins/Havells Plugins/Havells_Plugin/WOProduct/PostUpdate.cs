using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.WOProduct
{
    public class PostUpdate : IPlugin
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
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderproduct.EntityLogicalName.ToLower() && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorderproduct iProductWo = entity.ToEntity<msdyn_workorderproduct>();
                    Entity preImage = (Entity)context.PreEntityImages["image"];
                    tracingService.Trace("1");
                    {
                        EntityReference iReplacedPart = new EntityReference();
                        //Check part availablity on change of spare part and update inventory
                        Custom_Availablity tempCustom_Availablity = UpdateAvailiblityStatus(entity, preImage, service, tracingService);
                        UpdateInventory(entity, preImage, tempCustom_Availablity, service, tracingService);
                        if (entity.Attributes.Contains("hil_replacedpart"))
                        {
                            tracingService.Trace("POPULATE AMOUNT - 1");
                            if (iProductWo.hil_replacedpart != null)
                            {
                                tracingService.Trace("POPULATE AMOUNT - 2");
                                iReplacedPart = entity.GetAttributeValue<EntityReference>("hil_replacedpart");
                                Product iRepPrice = (Product)service.Retrieve(Product.EntityLogicalName, iReplacedPart.Id, new ColumnSet("hil_amount"));
                                if (iRepPrice.hil_Amount != null)
                                {
                                    tracingService.Trace("POPULATE AMOUNT - 3");
                                    msdyn_workorderproduct iProd = new msdyn_workorderproduct();
                                    iProd.Id = entity.Id;
                                    iProd.hil_PartAmount = iRepPrice.hil_Amount.Value;
                                    service.Update(iProd);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WOProduct.PostUpdate.Execute" + ex.Message);
            }
        }
        public static Custom_Availablity UpdateAvailiblityStatus(Entity entity, Entity preImage, IOrganizationService service, ITracingService tracingService)
        {
            Boolean result = false;
            Custom_Availablity tempCustom_Availablity = new Custom_Availablity();
            Int32 ReqQuantity = 0;
            try
            {
                //tracingService.Trace("UpdateAvailiblityStatus 2");
                msdyn_workorderproduct enWOProduct = entity.ToEntity<msdyn_workorderproduct>();
                #region If Replaced Part is not Null
                if (enWOProduct.hil_replacedpart != null)
                {
                    //tracingService.Trace("UpdateAvailiblityStatus 3");
                    //UpdateAvailiblityStatus(enWOProduct, service);
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        //tracingService.Trace("UpdateAvailiblityStatus 4");
                        var obj = from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                                  join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>() on _woProduct.msdyn_WorkOrder.Id equals _WorkOrder.Id
                                  // join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.msdyn_ServiceAccount.Id equals _Account.Id
                                  join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.hil_OwnerAccount.Id equals _Account.Id
                                  join _Owner in orgContext.CreateQuery<SystemUser>() on _Account.OwnerId.Id equals _Owner.Id
                                  where _woProduct.msdyn_workorderproductId.Value == enWOProduct.msdyn_workorderproductId.Value
                                  select new
                                  {
                                      _woProduct.hil_replacedpart,
                                      _woProduct.msdyn_Quantity,
                                      _woProduct.hil_WarrantyStatus,
                                      _Account.AccountId,
                                      _Owner.SystemUserId,
                                      _Account.Name,
                                      _Owner.FullName
                                  };
                        foreach (var iobj in obj)
                        {
                            //tracingService.Trace("UpdateAvailiblityStatus 5" + iobj.FullName + " " + iobj.Name);
                            Guid fsPart = Guid.Empty;
                            Guid fsOwnerAccount = Guid.Empty;
                            Guid fsOwner = Guid.Empty;
                            OptionSetValue opInventoryType = null; //to do need value
                            Double dQuantityRequired = 0;

                            if (iobj.hil_replacedpart != null)
                            {
                                fsPart = iobj.hil_replacedpart.Id;
                            }
                            if (iobj.SystemUserId != null)
                            {
                                fsOwner = iobj.SystemUserId.Value;
                            }
                            if (iobj.AccountId != null)
                            {
                                fsOwnerAccount = iobj.AccountId.Value;
                            }
                            if (iobj.msdyn_Quantity != null)
                            {
                                dQuantityRequired = iobj.msdyn_Quantity.Value;
                                ReqQuantity = Convert.ToInt32(iobj.msdyn_Quantity.Value);
                            }
                            else
                            {
                                ReqQuantity = 0;
                            }
                            if (iobj.hil_WarrantyStatus != null)
                            {
                                opInventoryType = iobj.hil_WarrantyStatus;
                                if (opInventoryType.Value == 3 || opInventoryType.Value == 4)
                                {
                                    opInventoryType = new OptionSetValue(2);
                                }
                            }
                            hil_inventory enInventory = HelperInventory.FindInventoryCommon(fsPart, fsOwnerAccount, fsOwner, service);
                            tracingService.Trace("UpdateAvailiblityStatus 6 Inv rec id" + enInventory.Id);
                            if (enInventory.Id != Guid.Empty)
                            {
                                if (enInventory.hil_AvailableQty != null && enInventory.hil_AvailableQty.Value >= dQuantityRequired)
                                {
                                    result = true;
                                }
                            }
                        }
                        #region UpdateAvailablityStatus

                        tracingService.Trace("UpdateAvailiblityStatus 7 Inv Availablity" + result);
                        msdyn_workorderproduct upWOProduct = new msdyn_workorderproduct();
                        upWOProduct.msdyn_workorderproductId = enWOProduct.msdyn_workorderproductId.Value;
                        if (result)
                        {
                            upWOProduct.hil_AvailabilityStatus = new OptionSetValue(1); //available
                            tempCustom_Availablity = Custom_Availablity.Available;
                            upWOProduct["hil_pendingquantity"] = Convert.ToInt32(0);
                        }
                        else
                        {
                            upWOProduct.hil_AvailabilityStatus = new OptionSetValue(2); // Not available
                            upWOProduct["hil_pendingquantity"] = Convert.ToInt32(ReqQuantity);
                            tempCustom_Availablity = Custom_Availablity.NonAvailable;
                        }
                        service.Update(upWOProduct);

                        #endregion

                    }
                }
                #endregion
                #region Replaced Part null
                else if (enWOProduct.hil_replacedpart == null ||
                    (preImage.Attributes.Contains("hil_replacedpart") && preImage.Contains("hil_replacedpart")
                    && preImage.GetAttributeValue<EntityReference>("hil_replacedpart").Id != enWOProduct.hil_replacedpart.Id)) //Now replaced part removed
                {
                        tracingService.Trace("UpdateAvailiblityStatus 7 Inv Availablity" + result);
                        msdyn_workorderproduct PreenWOProduct = preImage.ToEntity<msdyn_workorderproduct>();
                    if (PreenWOProduct.hil_replacedpart != null && enWOProduct.hil_replacedpart == null)
                    {
                        #region Update Availablity Status to Blank
                        msdyn_workorderproduct upWOProduct = new msdyn_workorderproduct();
                        upWOProduct.msdyn_workorderproductId = enWOProduct.msdyn_workorderproductId.Value;
                        //    upWOProduct.hil_AvailabilityStatus = null; //blank
                        service.Update(upWOProduct);
                        tempCustom_Availablity = Custom_Availablity.NA;
                        #endregion
                    }
                    else if(PreenWOProduct.hil_replacedpart != null && enWOProduct.hil_replacedpart != null)
                    {

                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WOProduct.PostUpdate.UpdateAvailiblityStatus" + ex.Message);
                //throw new InvalidPluginExecutionException(ex.Message);
            }
            return tempCustom_Availablity;
        }

        public static void UpdateInventory(Entity entity, Entity preImage, Custom_Availablity tempCustom_Availablity, IOrganizationService service, ITracingService tracing)
        {
            try
            {
                tracingService.Trace("UI1");
                msdyn_workorderproduct enWOProduct = entity.ToEntity<msdyn_workorderproduct>();
                #region on Selectin Replaced Part
                if (enWOProduct.hil_replacedpart != null)
                {
                    tracingService.Trace("UI2");
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        //enWOProduct = (msdyn_workorderproduct)service.Retrieve(msdyn_workorderproduct.EntityLogicalName, enWOProduct.Id, new ColumnSet(true));
                        tracingService.Trace("UI3" + enWOProduct.hil_AvailabilityStatus);
                        //if (enWOProduct.hil_AvailabilityStatus.Value == ((int)msdyn_workorderproduct_hil_AvailabilityStatus.Available))
                        {

                            tracingService.Trace("UI4");
                            var obj = from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                                      join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>() on _woProduct.msdyn_WorkOrder.Id equals _WorkOrder.Id
                                      //  join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.msdyn_ServiceAccount.Id equals _Account.Id
                                      join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.hil_OwnerAccount.Id equals _Account.Id
                                      join _Owner in orgContext.CreateQuery<SystemUser>() on _Account.OwnerId.Id equals _Owner.Id
                                      where _woProduct.msdyn_workorderproductId.Value == enWOProduct.msdyn_workorderproductId.Value
                                      select new
                                      {
                                          _woProduct.hil_replacedpart,
                                          _woProduct.msdyn_Quantity,
                                          _woProduct.hil_AvailabilityStatus,
                                          _woProduct.hil_WarrantyStatus,
                                          _Account.AccountId,
                                          _Owner.SystemUserId,
                                          _Account.Name,
                                          _Owner.FullName
                                      };

                            foreach (var iobj in obj)
                            {
                                tracingService.Trace("UI5 account name " + iobj.Name + " Owner " + iobj.FullName);
                                Guid fsPart = Guid.Empty;
                                Guid fsOwnerAccount = Guid.Empty;
                                Guid fsOwner = Guid.Empty;
                                OptionSetValue opInventoryType = null; //to do need value
                                Int32 iQuantityRequired = 0;
                                Guid fsWOProductId = enWOProduct.Id;

                                if (iobj.hil_replacedpart != null)
                                {
                                    fsPart = iobj.hil_replacedpart.Id;
                                }
                                if (iobj.SystemUserId != null)
                                {
                                    fsOwner = iobj.SystemUserId.Value;
                                }
                                if (iobj.AccountId != null)
                                {
                                    fsOwnerAccount = iobj.AccountId.Value;
                                }
                                if (iobj.msdyn_Quantity != null)
                                {
                                    iQuantityRequired = (Int32)iobj.msdyn_Quantity.Value;
                                }
                                if (iobj.hil_WarrantyStatus != null)
                                {
                                    opInventoryType = iobj.hil_WarrantyStatus;
                                    if (opInventoryType.Value == 3 || opInventoryType.Value == 4)
                                    {
                                        opInventoryType = new OptionSetValue(2);
                                    }
                                }
                                //Update inventory of current available item
                                if (tempCustom_Availablity == Custom_Availablity.Available)
                                {
                                    //HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, -iQuantityRequired, iQuantityRequired);
                                    HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, -iQuantityRequired, 0);
                                }
                                //change 22 08 2018
                                //if part not available 0 decrease in available 1 increase in defectinve
                                else if (tempCustom_Availablity == Custom_Availablity.NonAvailable)
                                {
                                    //   HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, 0, iQuantityRequired);
                                }
                                #region For old replaced Part
                                msdyn_workorderproduct PreenWOProduct = preImage.ToEntity<msdyn_workorderproduct>();
                                if (PreenWOProduct.hil_replacedpart != null)
                                {
                                    fsPart = PreenWOProduct.hil_replacedpart.Id;
                                    //Increase good quantity and decrease defective qty
                                    //HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, iQuantityRequired, -iQuantityRequired);
                                    HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, iQuantityRequired, 0);
                                }
                                #endregion
                            }
                        }
                    }
                }
                #endregion
                #region on removing Replaced Part
                else if (enWOProduct.hil_replacedpart == null) //Now replaced part removed
                {
                    msdyn_workorderproduct PreenWOProduct = preImage.ToEntity<msdyn_workorderproduct>();
                    if (PreenWOProduct.hil_replacedpart != null)
                    {
                        tracingService.Trace("UI12");
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            //Check part availablity - If blank or available
                            enWOProduct = (msdyn_workorderproduct)service.Retrieve(msdyn_workorderproduct.EntityLogicalName, enWOProduct.Id, new ColumnSet(true));
                            tracingService.Trace("UI13" + enWOProduct.hil_AvailabilityStatus);
                            if (
                                enWOProduct.hil_AvailabilityStatus == null ||
                                (enWOProduct.hil_AvailabilityStatus.Value == ((int)msdyn_workorderproduct_hil_AvailabilityStatus.Available))
                              )
                            {
                                tracingService.Trace("UI14");
                                var obj = from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                                          join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>() on _woProduct.msdyn_WorkOrder.Id equals _WorkOrder.Id
                                          join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.hil_OwnerAccount.Id equals _Account.Id
                                          join _Owner in orgContext.CreateQuery<SystemUser>() on _Account.OwnerId.Id equals _Owner.Id
                                          where _woProduct.msdyn_workorderproductId.Value == enWOProduct.msdyn_workorderproductId.Value
                                          select new
                                          {
                                              _woProduct.hil_replacedpart,
                                              _woProduct.msdyn_Quantity,
                                              _woProduct.hil_WarrantyStatus,
                                              _Account.AccountId,
                                              _Owner.SystemUserId,
                                              _Account.Name,
                                              _Owner.FullName,
                                          };
                                foreach (var iobj in obj)
                                {
                                    tracingService.Trace("UI15 Owner:" + iobj.FullName + " Account: " + iobj.Name);
                                    Guid fsPart = Guid.Empty;
                                    Guid fsOwnerAccount = Guid.Empty;
                                    Guid fsOwner = Guid.Empty;
                                    OptionSetValue opInventoryType = null; //to do need value
                                    Int32 iQuantityRequired = 0;
                                    Guid fsWOProductId = enWOProduct.Id;

                                    //Get part from pre entity
                                    if (PreenWOProduct.hil_replacedpart != null)
                                    {
                                        fsPart = PreenWOProduct.hil_replacedpart.Id;
                                    }
                                    if (iobj.SystemUserId != null)
                                    {
                                        fsOwner = iobj.SystemUserId.Value;
                                    }
                                    if (iobj.AccountId != null)
                                    {
                                        fsOwnerAccount = iobj.AccountId.Value;
                                    }
                                    if (iobj.msdyn_Quantity != null)
                                    {
                                        iQuantityRequired = (Int32)iobj.msdyn_Quantity.Value;
                                    }
                                    if (iobj.hil_WarrantyStatus != null)
                                        opInventoryType = iobj.hil_WarrantyStatus;
                                    tracingService.Trace("UI16");
                                    //change 958
                                    //if part wasn't available earlier, available  no change defective decrease
                                    if (PreenWOProduct.hil_AvailabilityStatus.Value == ((int)msdyn_workorderproduct_hil_AvailabilityStatus.NotAvailable))
                                    {
                                        tracingService.Trace("UI17");
                                        //HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, 0, -iQuantityRequired);
                                        //HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, 0, 0);
                                    }
                                    else
                                    {
                                        tracingService.Trace("UI18");
                                        //Increase good quantity and decrease defective qty
                                        //HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, iQuantityRequired, -iQuantityRequired);
                                        HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracing, iQuantityRequired, 0);
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WOProduct.PostUpdate.UpdateInventory" + ex.Message);
            }
        }
    }
}