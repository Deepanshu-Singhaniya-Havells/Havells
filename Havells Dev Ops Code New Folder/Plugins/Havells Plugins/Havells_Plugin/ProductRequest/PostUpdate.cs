using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
namespace Havells_Plugin.ProductRequest
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_productrequest.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    hil_productrequest ProductReq = (hil_productrequest)context.InputParameters["Target"];
                    //on quantity field chane
                    MSLQuantityValidation(ProductReq,service, context);
                    //on Fullfill boolean change
                    ManualPOFulfill(ProductReq,service, tracingService);
                    LevelOneApprover(ProductReq, service);
                    LevelTwoApprover(ProductReq, service);
                    LevelThreeApprover(ProductReq, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PostUpdate.Execute" + ex.Message);
            }
        }
        //Create Grn line when super franchisee or DSC fulfill the PO
        #region Manual PO Fulfil
        private static void ManualPOFulfill(hil_productrequest PO,IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    if (PO.hil_FulfIll.Value == true && PO.hil_FulfilledQuantity != null && PO.hil_FulfilledQuantity.Value > 0)
                    {
                        var obj = from _po in orgContext.CreateQuery<hil_productrequest>()
                                  join _owner in orgContext.CreateQuery<SystemUser>() on _po.OwnerId.Id equals _owner.Id
                                  join _account in orgContext.CreateQuery<Account>() on _po.hil_SuperFranchiseeDSEName.Id equals _account.Id
                                  where _po.Id == PO.Id
                                  select new
                                  {
                                      _po.hil_PartCode,
                                      _po.hil_SuperFranchiseeDSEName,
                                      _po.OwnerId,
                                      _po.hil_FulfilledQuantity,
                                      _po.hil_Quantity,
                                      _po.hil_WarrantyStatus,
                                      _owner.hil_Account,
                                      _POAccountOwner=_account.OwnerId
                                  };
                        foreach (var iobj in obj)
                        {
                            #region Initialise Variables
                            Guid fsPart = Guid.Empty;
                            Guid fsAccount = Guid.Empty;
                            Guid fsOwner = Guid.Empty;
                            Guid fsPOAcountOwner = Guid.Empty; 

                            OptionSetValue opInventoryType = null; //to do need value
                            Int32 iQuantity = 0;
                            
                            if (iobj.hil_PartCode != null)
                            {
                                fsPart = iobj.hil_PartCode.Id;
                            }
                            if (iobj.OwnerId != null)
                            {
                                fsOwner = iobj.OwnerId.Id;
                            }
                            if (iobj.hil_Account != null)
                            {
                                fsAccount = iobj.hil_Account.Id;
                            }
                            if (iobj._POAccountOwner != null)
                            {
                                fsPOAcountOwner = iobj._POAccountOwner.Id;
                            }
                            if (iobj.hil_FulfilledQuantity != null)
                            {
                                iQuantity = iobj.hil_FulfilledQuantity.Value;
                            }
                            if (iobj.hil_WarrantyStatus != null)
                            {
                                opInventoryType = iobj.hil_WarrantyStatus;
                            }
                            #endregion
                            hil_inventory enInventory = HelperInventory.FindInventoryCommon(fsPart, fsAccount, fsOwner, service);
                            if (enInventory.Id != Guid.Empty)
                            {
                                if (enInventory.hil_AvailableQty != null && enInventory.hil_AvailableQty.Value < iQuantity)
                                {
                                    throw new InvalidPluginExecutionException("Can't process the PO as available quantity is less than fullfilled quantity. Fullfill quantity-" + iQuantity + " Available Quantity-" + enInventory.hil_AvailableQty.Value);
                                }
                                else if (enInventory.hil_AvailableQty != null && enInventory.hil_AvailableQty.Value >= iQuantity)
                                {
                                    //reduce the Inventory of Super frenchise
                                    HelperInvJournal.CreateInvJournal(fsPart, fsAccount, fsOwner, Guid.Empty, Guid.Empty, Guid.Empty,PO.Id, Guid.Empty, service, iTrace , - iQuantity);
                                    //Create GRN Line and 
                                    hil_grnline crGrnLine = new hil_grnline();
                                        crGrnLine.hil_Quantity = iQuantity;
                                    if (fsPart != Guid.Empty)
                                        crGrnLine.hil_ProductCode = new EntityReference(Product.EntityLogicalName, fsPart);
                                    if (fsAccount != Guid.Empty)
                                    {
                                        crGrnLine.hil_Account = new EntityReference(Account.EntityLogicalName, fsAccount);
                                    }
                                    if (fsPOAcountOwner != Guid.Empty)
                                    {
                                        crGrnLine.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsPOAcountOwner);
                                    }
                                    crGrnLine.hil_ProductRequest = new EntityReference(PO.LogicalName,PO.Id);
                                    service.Create(crGrnLine);

                                }

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PostUpdate.ManualPOFulfill" + ex.Message);
            }
        }
        #endregion
        #region MSL Quantity Validation
        private static void MSLQuantityValidation(hil_productrequest ProductReq,IOrganizationService service, IPluginExecutionContext context)
        {
            try
            {
                if (ProductReq.hil_Quantity != null)
                {
                        if (ProductReq.hil_PRType.Equals(hil_PRType.MSL))
                        {
                        hil_productrequest PreImagePdt = (hil_productrequest)context.PreEntityImages["WoPdtImage"];
                        int NewQty = (int)ProductReq.hil_Quantity;
                        if (PreImagePdt.Attributes.Contains("msdyn_quantity"))
                        {
                            int OldQty = (int)PreImagePdt.hil_Quantity;
                            if (NewQty < OldQty)
                            {
                                throw new InvalidPluginExecutionException("Quantity can't be lesser than required MSL quantity");
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PostUpdate.CheckQuantityValidation" + ex.Message);
            }
        }
        #endregion
        #region Level One Approver
        public static void LevelOneApprover(hil_productrequest Req, IOrganizationService service)
        {
            hil_productrequest ProdReq = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, Req.Id, new ColumnSet("hil_division"));
            int NextLevel = 2;
            if (Req.hil_level1status != null)
            {
                if(Req.hil_level1status.Value == 1 && ProdReq.hil_Division != null)//Approved
                {
                    AssignToNextLevel(service, Req.hil_Division.Id, NextLevel, Req.Id);
                }
                else if(Req.hil_level1status.Value == 2)//Rejected
                {
                    Req.statuscode = new OptionSetValue(1);
                    if (ProdReq.hil_Job != null)
                    {
                        msdyn_workorder Ticket = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, ProdReq.hil_Job.Id, new ColumnSet(false));
                        Guid SubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Service Initiated", service);
                        if (SubStatus != Guid.Empty)
                        {
                            Ticket.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, SubStatus);
                            service.Update(Ticket);
                        }
                    }
                    service.Update(Req);
                }
            }
        }
        #endregion
        #region Level Two Approver
        public static void LevelTwoApprover(hil_productrequest Req, IOrganizationService service)
        {
            hil_productrequest ProdReq = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, Req.Id, new ColumnSet("hil_division"));
            int NextLevel = 3;
            if (Req.hil_level2status != null)
            {
                if (Req.hil_level2status.Value == 1)
                {
                    AssignToNextLevel(service, Req.hil_Division.Id, NextLevel, Req.Id);
                }
                else if (Req.hil_level2status.Value == 2)
                {
                    Req.statuscode = new OptionSetValue(1);
                    if (ProdReq.hil_Job != null)
                    {
                        msdyn_workorder Ticket = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, ProdReq.hil_Job.Id, new ColumnSet(false));
                        Guid SubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Service Initiated", service);
                        if (SubStatus != Guid.Empty)
                        {
                            Ticket.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, SubStatus);
                            service.Update(Ticket);
                        }
                    }
                    service.Update(Req);
                }
            }
        }
        #endregion
        #region Level Three Approver
        public static void LevelThreeApprover(hil_productrequest Req, IOrganizationService service)
        {
            hil_productrequest ProdReq = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, Req.Id, new ColumnSet("hil_workorderincident"));
            if (Req.hil_level3status != null)
            {
                if (Req.hil_level3status.Value == 1)
                {
                    Req.statuscode = new OptionSetValue(910590003);
                    if(ProdReq.hil_Job != null)
                    {
                        msdyn_workorder Ticket = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, ProdReq.hil_Job.Id, new ColumnSet(false));
                        Guid SubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Product Replacement in Progress", service);
                        if (SubStatus != Guid.Empty)
                        {
                            Ticket.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, SubStatus);
                            service.Update(Ticket);
                        }
                    }
                }
                else if (Req.hil_level3status.Value == 2)
                {
                    Req.statuscode = new OptionSetValue(1);
                    if (ProdReq.hil_Job != null)
                    {
                        msdyn_workorder Ticket = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, ProdReq.hil_Job.Id, new ColumnSet(false));
                        Guid SubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Service Initiated", service);
                        if (SubStatus != Guid.Empty)
                        {
                            Ticket.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, SubStatus);
                            service.Update(Ticket);
                        }
                    }
                }
                service.Update(Req);
            }
        }
        #endregion
        #region Conditional Assign
        public static void AssignToNextLevel(IOrganizationService service, Guid Division, int NextLevel, Guid ProdReq)
        {
            QueryByAttribute Query = new QueryByAttribute(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.AddAttributeValue("hil_division", Division);
            Query.AddAttributeValue("hil_levelofapprover", NextLevel);
            Query.AddAttributeValue("new_warrantystatus", 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count >= 1)
            {
                foreach(hil_integrationconfiguration Conf in Found.Entities)
                {
                    if(Conf.hil_approvername != null)
                    {
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, ProdReq, service);
                    }
                }
            }
        }
        #endregion
    }
}