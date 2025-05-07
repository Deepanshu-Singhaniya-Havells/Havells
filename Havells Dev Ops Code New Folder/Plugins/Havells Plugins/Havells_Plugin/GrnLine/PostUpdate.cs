using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.GrnLine
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
                    && context.PrimaryEntityName.ToLower() == hil_grnline.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    //When Grn Line is confirmed by user at their end after putting quantity
                    onGrnLineConfirm(entity, service, tracingService);

                    //On approval of defecetive/missing qty by BSH, update inventory, assign back grn line and create Return Line
                    GrnLineApprovebyBSH(entity, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.Execute" + ex.Message);
            }
        }
        #region GRNLineConfirm
        public static void onGrnLineConfirm(Entity entity, IOrganizationService service, ITracingService iTrace)
        {
            if (entity.Contains("statuscode") && entity.GetAttributeValue<OptionSetValue>("statuscode").Value == 910590000)//Confirmed
            {
                #region Initialize Variables
                hil_grnline iGrn = (hil_grnline)service.Retrieve(hil_grnline.EntityLogicalName, entity.Id, new ColumnSet(true));
                Int32 iRequired = 0;
                Int32 iPending = 0;
                Int32 Fulfilled = 0;
                Int32 iRemaining = 0;
                Guid iAccount = iGrn.hil_Account.Id;
                Account iFranch = (Account)service.Retrieve(Account.EntityLogicalName, iAccount, new ColumnSet("ownerid"));
                Guid iAccountOwner = iFranch.OwnerId.Id;
                Guid iPart = iGrn.hil_ProductCode.Id;
                Guid iGrnId = iGrn.Id;
                OptionSetValue iInventoryType = iGrn.hil_WarrantyStatus;
                Guid iWoProduct = new Guid();
                #endregion
                QuantityValidationOwner(entity, service);
                #region Update Good Quantity
                if(iGrn.hil_ProductRequest != null)
                {
                    Fulfilled = Convert.ToInt32(iGrn.hil_GoodQuantity);
                    hil_productrequest iProdReq = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, iGrn.hil_ProductRequest.Id, new ColumnSet("hil_prtype", "hil_job"));
                    if(iProdReq.hil_PRType != null && iProdReq.hil_PRType.Value == 910590003 && 
                        iProdReq.hil_Job != null)
                    {
                        QueryExpression Query = new QueryExpression(msdyn_workorderproduct.EntityLogicalName);
                        Query.ColumnSet = new ColumnSet(true);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, iProdReq.hil_Job.Id);
                        Query.Criteria.AddCondition("hil_replacedpart", ConditionOperator.Equal, iGrn.hil_ProductCode.Id);
                        Query.Criteria.AddCondition("hil_availabilitystatus", ConditionOperator.Equal, 2);//Not Available
                        Query.Criteria.AddCondition("hil_linestatus", ConditionOperator.Equal, 1);//Part Requested
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        if(Found.Entities.Count == 1)
                        {
                            iWoProduct = Found.Entities[0].Id;
                            msdyn_workorderproduct iJPro = Found.Entities[0].ToEntity<msdyn_workorderproduct>();
                            //iJPro.hil_replacedpart
                            if(iJPro.msdyn_Quantity != null)
                            {
                                iRequired = Convert.ToInt32(iJPro.msdyn_Quantity);
                                if(iJPro.Attributes.Contains("hil_pendingquantity"))
                                {
                                    iPending = Convert.ToInt32(iJPro["hil_pendingquantity"]);
                                    if(iPending < Fulfilled)
                                    {
                                        iRemaining = Fulfilled - iPending;
                                        msdyn_workorderproduct iJUpPro = new msdyn_workorderproduct();
                                        iJUpPro.Id = iJPro.Id;
                                        iJUpPro.hil_AvailabilityStatus = new OptionSetValue(1);
                                        iJUpPro.hil_LineStatus = new OptionSetValue(910590000);
                                        iJUpPro["hil_markused"] = true;
                                        iJUpPro["hil_pendingquantity"] = 0;
                                        service.Update(iJUpPro);
                                        HelperInvJournal.CreateInvJournal(iPart, iAccount, iAccountOwner, iWoProduct, iGrnId, Guid.Empty, Guid.Empty, Guid.Empty, service, iTrace, iRemaining);
                                    }
                                    else if(iPending == Fulfilled)
                                    {
                                        msdyn_workorderproduct iJUpPro = new msdyn_workorderproduct();
                                        iJUpPro.Id = iJPro.Id;
                                        iJUpPro.hil_AvailabilityStatus = new OptionSetValue(1);
                                        iJUpPro.hil_LineStatus = new OptionSetValue(910590000);
                                        iJUpPro["hil_markused"] = true;
                                        iJUpPro["hil_pendingquantity"] = 0;
                                        service.Update(iJUpPro);
                                    }
                                    else if(iPending > Fulfilled)
                                    {
                                        msdyn_workorderproduct iJUpPro = new msdyn_workorderproduct();
                                        iJUpPro.Id = iJPro.Id;
                                        iJUpPro["hil_pendingquantity"] = iPending - Fulfilled;
                                        service.Update(iJUpPro);
                                    }
                                }
                            }
                        }
                        else if(Found.Entities.Count > 1)
                        {
                            iRemaining = Fulfilled;
                            foreach (msdyn_workorderproduct iJPro in Found.Entities)
                            {
                                if (iJPro.Attributes.Contains("hil_pendingquantity"))
                                {
                                    iPending = Convert.ToInt32(iJPro["hil_pendingquantity"]);
                                    if (iPending < iRemaining)
                                    {
                                        iRemaining = iRemaining - iPending;
                                        msdyn_workorderproduct iJUpPro = new msdyn_workorderproduct();
                                        iJUpPro.Id = iJPro.Id;
                                        iJUpPro.hil_AvailabilityStatus = new OptionSetValue(1);
                                        iJUpPro.hil_LineStatus = new OptionSetValue(910590000);
                                        iJUpPro["hil_markused"] = true;
                                        iJUpPro["hil_pendingquantity"] = 0;
                                        service.Update(iJUpPro);
                                    }
                                    else if (iPending == iRemaining)
                                    {
                                        msdyn_workorderproduct iJUpPro = new msdyn_workorderproduct();
                                        iJUpPro.Id = iJPro.Id;
                                        iJUpPro.hil_AvailabilityStatus = new OptionSetValue(1);
                                        iJUpPro.hil_LineStatus = new OptionSetValue(910590000);
                                        iJUpPro["hil_markused"] = true;
                                        iJUpPro["hil_pendingquantity"] = 0;
                                        service.Update(iJUpPro);
                                    }
                                    else if (iPending > iRemaining)
                                    {
                                        msdyn_workorderproduct iJUpPro = new msdyn_workorderproduct();
                                        iJUpPro.Id = iJPro.Id;
                                        iJUpPro["hil_pendingquantity"] = iPending - iRemaining;
                                        service.Update(iJUpPro);
                                    }
                                }
                            }
                            if(iRemaining > 0)
                            {
                                HelperInvJournal.CreateInvJournal(iPart, iAccount, iAccountOwner, Guid.Empty, iGrnId, Guid.Empty, Guid.Empty, Guid.Empty, service, iTrace, iRemaining);
                            }
                        }
                        else
                        {
                            UpdateGoodQty(entity, service, iTrace);
                        }
                    }
                    else
                    {
                        UpdateGoodQty(entity, service, iTrace);
                    }
                }
                #endregion
                AssignGrnForApproval(entity,service);
            }
        }
        public static void AssignGrnForApproval(Entity entity, IOrganizationService service)
        {
            try
            {
                hil_grnline GrnLine = entity.ToEntity<hil_grnline>();
                //if defective or missing qty exist
                if ((GrnLine.hil_MissingQuantity != null && GrnLine.hil_MissingQuantity.Value > 0) ||
                    (GrnLine.hil_DefectiveQuantity != null && GrnLine.hil_DefectiveQuantity.Value > 0))
                {
                    //assign to ERD
                    Guid fsBranchHead = GetBranchHead(GrnLine.hil_Account.Id,service);
                    if (fsBranchHead != Guid.Empty)
                    {
                        Account enAccount = (Account)service.Retrieve(Account.EntityLogicalName, fsBranchHead, new ColumnSet("ownerid"));
                        if (enAccount.OwnerId != null)
                        {
                            Helper.Assign(SystemUser.EntityLogicalName, hil_grnline.EntityLogicalName, enAccount.OwnerId.Id, GrnLine.Id, service);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.AssignGrnForApproval" + ex.Message);
            }
         }
        public static Guid GetBranchHead(Guid fsAccountId, IOrganizationService service)
        {
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                Guid fsResult = Guid.Empty;
                var obj = from _Account in orgContext.CreateQuery<Account>()
                          where _Account.Id == fsAccountId
                          select new {
                              _Account.CustomerTypeCode,
                          _Account.Id,
                          _Account.OwnerId
                          };
                foreach (var iobj in obj)
                {
                    if (iobj.CustomerTypeCode != null && iobj.CustomerTypeCode.Equals(Account_CustomerTypeCode.Branch))
                    {
                        return iobj.Id;
                    }
                    else
                    {
                        fsResult = GetBranchHead(iobj.Id, service);
                    }
                }
                return fsResult;
             }
        }
        public static void UpdateGoodQty(Entity entity, IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                hil_grnline GrnLine = (hil_grnline)service.Retrieve(hil_grnline.EntityLogicalName, entity.Id, new ColumnSet(true));// entity.ToEntity<hil_grnline>();

                Guid fsPart = Guid.Empty;
                Guid fsOwnerAccount = Guid.Empty;
                Guid fsOwner = Guid.Empty;
                OptionSetValue opInventoryType = null;
                Int32 iGoodQty = 0;
                Guid fsWOProductId = Guid.Empty;
                Guid fsGRNLines = Guid.Empty;

                if (GrnLine.hil_ProductCode != null)
                    fsPart = GrnLine.hil_ProductCode.Id;
                if (GrnLine.hil_Account != null)
                    fsOwnerAccount= GrnLine.hil_Account.Id;
                if (GrnLine.OwnerId!= null)
                    fsOwner = GrnLine.OwnerId.Id;
                if (GrnLine.hil_WarrantyStatus != null)
                    opInventoryType = GrnLine.hil_WarrantyStatus;
                if (GrnLine.hil_GoodQuantity != null)
                    iGoodQty = GrnLine.hil_GoodQuantity.Value;
                fsGRNLines = GrnLine.Id;

                HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, fsWOProductId, fsGRNLines,Guid.Empty ,Guid.Empty, Guid.Empty,service, iTrace, iGoodQty);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.UpdateGoodQty" + ex.Message);
            }
        }
        public static void QuantityValidationOwner(Entity entity,IOrganizationService service)
        {
            try
            {
                Int32 iTotalQty = 0;
                Int32 iGoodQty = 0;
                Int32 iDefQty = 0;
                Int32 iMissQty = 0;

                hil_grnline  GrnLine = (hil_grnline)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                if (GrnLine.hil_Quantity != null)
                {
                    iTotalQty = GrnLine.hil_Quantity.Value;
                }
                if (GrnLine.hil_GoodQuantity != null)
                {
                    iGoodQty = GrnLine.hil_GoodQuantity.Value;
                }
                if (GrnLine.hil_DefectiveQuantity != null)
                {
                    iDefQty = GrnLine.hil_DefectiveQuantity.Value;
                }
                if (GrnLine.hil_MissingQuantity != null)
                {
                    iMissQty = GrnLine.hil_MissingQuantity.Value;
                }
                if (iTotalQty != (iGoodQty + iDefQty + iMissQty))
                {
                    throw new InvalidPluginExecutionException("Sum of Good, defective and missing quantity should be equal to total quantity.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.QuantityValidation" + ex.Message);
            }

        }
        #endregion
        #region GRNLineApproveBSH
        public static void GrnLineApprovebyBSH(Entity entity, IOrganizationService servic, ITracingService iTrace)
        {
            try
            {
                hil_grnline GrnLine = entity.ToEntity<hil_grnline>();
                if (GrnLine.hil_ApproverStatus != null && GrnLine.hil_ApproverStatus.Value ==1)//apprver staus=done
                {
                    QuantityValidationApprover(entity,servic);

                    UpdateDefMissingQty(entity,servic, iTrace);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.GrfLineApprovebyBSH" + ex.Message);
            }


        }

        public static void UpdateDefMissingQty(Entity entity, IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                Guid fsPart = Guid.Empty;
                Guid fsGrnLineAccount = Guid.Empty;
                Guid fsOwner = Guid.Empty;
                OptionSetValue opInventoryType = null;
                Guid fsGRNLines = Guid.Empty;
                Guid fsGrnLineAccountOwner = Guid.Empty;
                Int32 iApproverDefQty = 0;
                Int32 iApproverMissQty = 0;
                Int32 iDefQty = 0;
                Int32 iMissQty = 0;
                Int32 iGoodQty = 0;

                #region GetGrnLineAccount'sOwner
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _GrnLine in orgContext.CreateQuery<hil_grnline>()
                              join _Account in orgContext.CreateQuery<Account>() on _GrnLine.hil_Account.Id equals _Account.Id
                              select new { _Account.OwnerId };
                    foreach (var iobj in obj)
                    {
                        fsGrnLineAccountOwner = iobj.OwnerId.Id;
                    }
                }
                #endregion

                hil_grnline GrnLine = (hil_grnline)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));

                if (GrnLine.hil_ApprovedDefectiveQuantity != null)
                {
                    iApproverDefQty = GrnLine.hil_ApprovedDefectiveQuantity.Value;
                }
                if (GrnLine.hil_ApprovedMissingQuantity != null)
                {
                    iApproverMissQty = GrnLine.hil_ApprovedMissingQuantity.Value;
                }
                if (GrnLine.hil_DefectiveQuantity != null)
                {
                    iDefQty = GrnLine.hil_DefectiveQuantity.Value;
                }
                if (GrnLine.hil_MissingQuantity != null)
                {
                    iMissQty = GrnLine.hil_MissingQuantity.Value;
                }
                iGoodQty = iDefQty - iApproverDefQty;
                iGoodQty += iMissQty - iApproverMissQty;
                
                if (GrnLine.hil_ProductCode != null)
                    fsPart = GrnLine.hil_ProductCode.Id;
                if (GrnLine.hil_Account != null)
                    fsGrnLineAccount = GrnLine.hil_Account.Id;
                if (GrnLine.OwnerId != null)
                    fsOwner = GrnLine.OwnerId.Id;
                if (GrnLine.hil_WarrantyStatus != null)
                    opInventoryType = GrnLine.hil_WarrantyStatus;
                if (GrnLine.hil_GoodQuantity != null)
                    iGoodQty = GrnLine.hil_GoodQuantity.Value;
                fsGRNLines = GrnLine.Id;

                //Create inv journal for grn line's account
                HelperInvJournal.CreateInvJournal(fsPart, fsGrnLineAccount, fsGrnLineAccountOwner, Guid.Empty, fsGRNLines, Guid.Empty, Guid.Empty, Guid.Empty, service, iTrace, iGoodQty);
                if (iApproverDefQty > 0)
                {
                    HelperInvJournal.CreateInvJournalDefective(fsPart, fsGrnLineAccount, fsGrnLineAccountOwner, opInventoryType, Guid.Empty, fsGRNLines, Guid.Empty, Guid.Empty, Guid.Empty, service, iTrace, iApproverDefQty);
                }

                //Assign Grn Line back to account's owner
                Helper.Assign(SystemUser.EntityLogicalName, hil_grnline.EntityLogicalName, fsGrnLineAccountOwner, GrnLine.Id, service);

                #region Create Return Header and Return Line for Missing Qty
                if (iMissQty > 0)
                {
                    hil_returnheader crRHeader = new hil_returnheader();
                    crRHeader.hil_Account = new EntityReference(Account.EntityLogicalName, fsGrnLineAccount);
                    crRHeader.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsGrnLineAccountOwner);
                    Guid fsRHeader = service.Create(crRHeader);

                    hil_ReturnLine crRLine = new hil_ReturnLine();
                    crRLine.hil_Account = new EntityReference(Account.EntityLogicalName, fsGrnLineAccount);
                    crRLine.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsGrnLineAccountOwner);
                    crRLine.hil_ProductCode = new EntityReference(Product.EntityLogicalName,fsPart);
                    crRLine.hil_Quantity = iMissQty;
                    crRLine.hil_ReturnType = new OptionSetValue(4);//Missing return refund
                    crRLine.hil_ReturnHeader = new EntityReference(hil_returnheader.EntityLogicalName, fsRHeader);
                    service.Create(crRLine);

                }
                #endregion

               
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.UpdateGoodQty" + ex.Message);
            }
        }

        public static void QuantityValidationApprover(Entity entity, IOrganizationService service)
        {
            try
            {
                Int32 iApproverDefQty = 0;
                Int32 iApproverMissQty = 0;
                Int32 iDefQty = 0;
                Int32 iMissQty = 0;

                hil_grnline GrnLine = (hil_grnline)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                if (GrnLine.hil_ApprovedDefectiveQuantity != null)
                {
                    iApproverDefQty = GrnLine.hil_ApprovedDefectiveQuantity.Value;
                }
                if (GrnLine.hil_ApprovedMissingQuantity != null)
                {
                    iApproverMissQty = GrnLine.hil_ApprovedMissingQuantity.Value;
                }
                if (GrnLine.hil_DefectiveQuantity != null)
                {
                    iDefQty = GrnLine.hil_DefectiveQuantity.Value;
                }
                if (GrnLine.hil_MissingQuantity != null)
                {
                    iMissQty = GrnLine.hil_MissingQuantity.Value;
                }
                if (iApproverDefQty> iDefQty)
                {
                    throw new InvalidPluginExecutionException("Approved defective quantity can not be greater than defective quantity.");
                }
                if (iApproverMissQty > iMissQty)
                {
                    throw new InvalidPluginExecutionException("Approved missing quantity can not be greater than missing quantity.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.GrnLine.PostUpdate.QuantityValidation" + ex.Message);
            }

        }
        #endregion
    }
}
