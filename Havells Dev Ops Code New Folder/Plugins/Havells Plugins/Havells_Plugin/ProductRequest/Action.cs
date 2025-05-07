using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;

namespace Havells_Plugin.ProductRequest
{
    public class Action : IPlugin
    {
       public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
             tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.MessageName== "hil_CheckStockinHandonPO")
                {
                    EntityReference enRef = (EntityReference)context.InputParameters["Target"];
                   //throw new InvalidPluginExecutionException("Test");
                    CheckStockInHand(enRef.Id,service);
                }
            } catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
}

        public static void CheckStockInHand(Guid fsPoId, IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _po in orgContext.CreateQuery<hil_productrequest>()
                              join _owner in orgContext.CreateQuery<SystemUser>() on _po.OwnerId.Id equals _owner.Id
                              join _account in orgContext.CreateQuery<Account>() on _po.hil_SuperFranchiseeDSEName.Id equals _account.Id
                              where _po.hil_productrequestId == fsPoId
                              select new
                              {
                                  _po.hil_PartCode,
                                  _po.hil_SuperFranchiseeDSEName,
                                  _po.OwnerId,
                                  _po.hil_FulfilledQuantity,
                                  _po.hil_Quantity,
                                  _po.hil_WarrantyStatus,
                                  _owner.hil_Account,
                                  _account.OwnerId.Id
                              };
                    foreach (var iobj in obj)
                    {
                        tracingService.Trace("1");

                        Guid fsPart = Guid.Empty;
                        Guid fsOwnerAccount = Guid.Empty;
                        Guid fsOwner = Guid.Empty;
                        OptionSetValue opInventoryType = null; //to do need value
                        Double dQuantityRequired = 0;

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
                            fsOwnerAccount = iobj.hil_Account.Id;
                        }
                        if (iobj.hil_FulfilledQuantity != null)
                        {
                            dQuantityRequired = iobj.hil_FulfilledQuantity.Value;
                        }
                        if (iobj.hil_WarrantyStatus != null)
                        {
                            opInventoryType = iobj.hil_WarrantyStatus;
                        }
                        tracingService.Trace("2");

                        hil_inventory enInventory = HelperInventory.FindInventory(fsPart, fsOwnerAccount, fsOwner, opInventoryType, service);
                        tracingService.Trace("3 Inventory Id"+ enInventory.Id);
                        if (enInventory.Id != Guid.Empty)
                        {
                            tracingService.Trace("4 Available qty" + enInventory.hil_AvailableQty.Value);
                            if (enInventory.hil_AvailableQty != null)
                            {
                                tracingService.Trace("5 "+ fsPoId);
                                hil_productrequest upProductReq = new hil_productrequest();
                                upProductReq.Id = fsPoId;
                                upProductReq.hil_hil_availablequantity = enInventory.hil_AvailableQty.Value;
                                service.Update(upProductReq);
                            }
                        }
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