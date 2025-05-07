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
    public class PostUpdateV1 : IPlugin
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
                msdyn_workorderproduct enUpdateWP = new msdyn_workorderproduct();
                enUpdateWP.Id = entity.Id;
                Entity PreEnJPro = (Entity)context.PreEntityImages["Preimage"];
                Entity PostEnJPro = (Entity)context.PostEntityImages["Postimage"];
                msdyn_workorderproduct preImageWOProduct = PreEnJPro.ToEntity<msdyn_workorderproduct>();
                msdyn_workorderproduct postImageWOProduct = PostEnJPro.ToEntity<msdyn_workorderproduct>();
                Guid accountid = Guid.Empty;
                Guid accountownerid = Guid.Empty;
                Int32 iQuantityRequired = new Int32();
                OutputClassOwner iOutput = new OutputClassOwner();
                iOutput = GetOwnerAccount(service, postImageWOProduct);
                if(iOutput.OwnerAccount != Guid.Empty)
                {
                    accountid = iOutput.OwnerAccount;
                    accountownerid = iOutput.Owner;
                }
                else
                {
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXX - FRANCHISE ACCOUNT BLANK IN JOB - XXXXXXXXXXXXXXXXXXXXXXXXXX");
                }
                if (postImageWOProduct.hil_Quantity == null || postImageWOProduct.hil_Quantity.Value == 0)
                {
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXX - QUANTITY CAN'T BE NULL OR ZERO - XXXXXXXXXXXXXXXXXXXXXXXXXX");
                }
                iQuantityRequired = Convert.ToInt32(postImageWOProduct.hil_Quantity.Value);
                if (preImageWOProduct.hil_replacedpart == null)
                {
                    if (entity.Contains("hil_replacedpart"))
                    {
                        //check inventory of current replaced part
                        hil_inventory enInventory = HelperInventory.FindInventoryCommon(entity.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                            accountid, accountownerid, service);
                        if (enInventory.Id != Guid.Empty || (enInventory.hil_AvailableQty != null && enInventory.hil_AvailableQty.Value >= postImageWOProduct.hil_Quantity.Value))
                        {

                            enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(1);//SET AVAILABLE
                            enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(0);
                            //add negative qty inventory journal of current replaced part if available
                            HelperInvJournal.CreateInvJournal(entity.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                                accountid, accountownerid, entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, -iQuantityRequired, 0);
                        }
                        else
                        {
                            enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(2);//SET NOT AVAILABLE
                            enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(iQuantityRequired);
                        }
                        service.Update(enUpdateWP);
                    }
                }
                else if (preImageWOProduct.hil_replacedpart != null && preImageWOProduct.hil_AvailabilityStatus.Value == 1)
                {
                    if (entity.Contains("hil_replacedpart"))
                    {
                        //check inventory of current replaced part
                        hil_inventory enInventory = HelperInventory.FindInventoryCommon(entity.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                            accountid, accountownerid, service);
                        if (enInventory.Id != Guid.Empty || (enInventory.hil_AvailableQty != null && enInventory.hil_AvailableQty.Value >= postImageWOProduct.hil_Quantity.Value))
                        {
                            enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(1);//SET AVAILABLE
                            enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(0);
                            HelperInvJournal.CreateInvJournal(entity.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                                accountid, accountownerid, entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, -iQuantityRequired, 0);
                            HelperInvJournal.CreateInvJournal(preImageWOProduct.hil_replacedpart.Id, accountid, accountownerid,
                                entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, iQuantityRequired, 0);
                            //add negative qty inventory journal of current replaced part if available
                        }
                        else
                        {
                            enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(2);//SET NOT AVAILABLE
                            enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(iQuantityRequired);
                            HelperInvJournal.CreateInvJournal(preImageWOProduct.hil_replacedpart.Id, accountid, accountownerid,
                                entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, iQuantityRequired, 0);
                        }
                        //add postive qty inventory journal of pre image replaced part
                    }
                    else
                    {
                        HelperInvJournal.CreateInvJournal(preImageWOProduct.hil_replacedpart.Id, accountid, accountownerid,
                                entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, iQuantityRequired, 0);
                        //add postive qty inventory journal of pre image replaced part
                    }
                    service.Update(enUpdateWP);
                }
                else if (preImageWOProduct.hil_replacedpart != null && preImageWOProduct.hil_AvailabilityStatus.Value == 2)
                {
                    if (entity.Contains("hil_replacedpart"))
                    {
                        //check inventory of current replaced part
                        hil_inventory enInventory = HelperInventory.FindInventoryCommon(entity.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                            accountid, accountownerid, service);
                        if (enInventory.Id != Guid.Empty || (enInventory.hil_AvailableQty != null && enInventory.hil_AvailableQty.Value >= postImageWOProduct.hil_Quantity.Value))
                        {
                            enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(1);//SET AVAILABLE
                            enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(0);
                            HelperInvJournal.CreateInvJournal(entity.GetAttributeValue<EntityReference>("hil_replacedpart").Id,
                                accountid, accountownerid, entity.Id, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService, -iQuantityRequired, 0);
                            //add negative qty inventory journal of current replaced part if available
                        }
                        else
                        {
                            enUpdateWP.hil_AvailabilityStatus = new OptionSetValue(2);//SET NOT AVAILABLE
                            enUpdateWP["hil_pendingquantity"] = Convert.ToInt32(iQuantityRequired);
                        }
                    }
                    else
                    {
                        //no action
                    }
                    service.Update(enUpdateWP);
                }
                else
                {

                }
            }
        }
        #region GET RELATED FRANCHISE AND OWNER
        public static OutputClassOwner GetOwnerAccount(IOrganizationService service, msdyn_workorderproduct enWOProduct)
        {
            OutputClassOwner iOutPut = new OutputClassOwner();
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
                    iOutPut.OwnerAccount = iobj.AccountId.Value;
                }
                return iOutPut;
            }
        }
        #endregion
    }
    public class OutputClassOwner
    {
        public Guid Owner { get; set; }
        public Guid OwnerAccount { get; set; }
    }
}