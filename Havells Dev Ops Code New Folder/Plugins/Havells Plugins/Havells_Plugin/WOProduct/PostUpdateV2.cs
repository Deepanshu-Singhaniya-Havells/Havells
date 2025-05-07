using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.WOProduct
{
    public class PostUpdateV2 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderproduct.EntityLogicalName.ToLower() && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                msdyn_workorderproduct postImageWOProduct = ((Entity)context.PostEntityImages["image"]).ToEntity<msdyn_workorderproduct>();

                //if (postImageWOProduct.hil_purchaseorder != null)
                //{
                //    throw new InvalidPluginExecutionException("PO has already been created for this Part. It cannot be updated.");
                //}
                #region Check Last Journal Posted is not null
                if (postImageWOProduct.Contains("hil_lastpostedjournal") && postImageWOProduct.GetAttributeValue<EntityReference>("hil_lastpostedjournal") != null)
                {
                    tracingService.Trace("Line 42 ");
                    //create reverse journal of last posted journal.
                    hil_inventoryjournal lastPostedJournal = (hil_inventoryjournal)service.Retrieve(hil_inventoryjournal.EntityLogicalName, postImageWOProduct.GetAttributeValue<EntityReference>("hil_lastpostedjournal").Id,
                        new ColumnSet(new string[] { "hil_part", "hil_owneraccount", "ownerid", "hil_availableqtychange" }));
                    tracingService.Trace("Line 45 ");

                    HelperInvJournal.CreateInvJournal(lastPostedJournal.hil_Part.Id, lastPostedJournal.hil_OwnerAccount.Id, lastPostedJournal.OwnerId.Id,
                        entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, -lastPostedJournal.hil_AvailableQtyChange.Value, 0);
                }
                #endregion
                tracingService.Trace("Line 52 ");
                #region check availability & create Journal if available

                if (postImageWOProduct.msdyn_Quantity != null && postImageWOProduct.msdyn_Quantity.Value > 0 && postImageWOProduct.hil_replacedpart != null)
                {
                    AccountOwnerCombo account_n_owner = GetOwnerAccount(service, postImageWOProduct);
                    Guid accountid = account_n_owner.Account != null ? account_n_owner.Account : throw new InvalidPluginExecutionException("Job Account is missing.");
                    Guid accountownerid = account_n_owner.Owner != null ? account_n_owner.Owner : throw new InvalidPluginExecutionException("Job Account Owner is missing");
                    tracingService.Trace("Line 57 ");
                    int iQuantityRequired = Convert.ToInt32(postImageWOProduct.msdyn_Quantity.Value);

                    hil_inventory availableInventory = HelperInventory.FindInventoryCommon(postImageWOProduct.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                                accountid, accountownerid, service);
                    if (availableInventory.Id != Guid.Empty && (availableInventory.hil_AvailableQty != null && availableInventory.hil_AvailableQty.Value >= iQuantityRequired))
                    {
                        //create Inventory Journal
                        //set created journal on lookup field "Last Posted Journal"
                        //set Availability Status as "Available"
                        msdyn_workorderproduct enUpdateWP = new msdyn_workorderproduct();
                        enUpdateWP.Id = entity.Id;
                        Guid journalId = HelperInvJournal.CreateInventoryJournal(postImageWOProduct.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                                        accountid, accountownerid, entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty,
                                        service, tracingService, -iQuantityRequired, 0);

                        enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(1);//SET AVAILABLE
                        enUpdateWP["hil_linestatus"] = new OptionSetValue(2);
                        enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(0);
                        enUpdateWP["hil_lastpostedjournal"] = new EntityReference(hil_inventoryjournal.EntityLogicalName, journalId);
                        service.Update(enUpdateWP);
                    }
                    else
                    {
                        //set Null on lookup field "Last Posted Journal"
                        //set Availability Status as "Not Available"
                        msdyn_workorderproduct enUpdateWP = new msdyn_workorderproduct();
                        enUpdateWP.Id = entity.Id;
                        enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(2);//SET NOT AVAILABLE
                        enUpdateWP["hil_linestatus"] = new OptionSetValue(1);
                        enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(iQuantityRequired);
                        enUpdateWP["hil_lastpostedjournal"] = null;
                        service.Update(enUpdateWP);
                    }
                }
                else
                {
                    msdyn_workorderproduct enUpdateWP = new msdyn_workorderproduct();
                    enUpdateWP.Id = entity.Id;
                    enUpdateWP.hil_AvailabilityStatus = null;
                    enUpdateWP.hil_replacedpart = null;
                    enUpdateWP["hil_lastpostedjournal"] = null;
                    enUpdateWP["hil_linestatus"] = null;
                    service.Update(enUpdateWP);
                }
                #endregion
            }
            else
            {
                msdyn_workorderproduct enUpdateWP = new msdyn_workorderproduct();
                enUpdateWP.Id = context.PrimaryEntityId;
                enUpdateWP.msdyn_Description = context.Depth.ToString();
                service.Update(enUpdateWP);
            }
        }
        #region GET RELATED FRANCHISE AND OWNER
        public static AccountOwnerCombo GetOwnerAccount(IOrganizationService service, msdyn_workorderproduct enWOProduct)
        {
            AccountOwnerCombo iOutPut = new AccountOwnerCombo();
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                var obj = (from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                           join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>() on _woProduct.msdyn_WorkOrder.Id equals _WorkOrder.Id
                           // join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.msdyn_ServiceAccount.Id equals _Account.Id
                           join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.hil_OwnerAccount.Id equals _Account.Id
                           join _Owner in orgContext.CreateQuery<SystemUser>() on _Account.OwnerId.Id equals _Owner.Id
                           where _woProduct.msdyn_workorderproductId.Value == enWOProduct.msdyn_workorderproductId.Value && _WorkOrder.hil_OwnerAccount != null
                           select new
                           {
                               _Account.AccountId,
                               _Owner.SystemUserId,
                               _Account.Name,
                               _Owner.FullName
                           }).Take(1);
                foreach (var iobj in obj)
                {
                    iOutPut.Owner = iobj.SystemUserId.Value;
                    iOutPut.Account = iobj.AccountId.Value;
                }
                return iOutPut;
            }
        }
        #endregion
    }
    public class AccountOwnerCombo
    {
        public Guid Owner { get; set; }
        public Guid Account { get; set; }
    }
}
