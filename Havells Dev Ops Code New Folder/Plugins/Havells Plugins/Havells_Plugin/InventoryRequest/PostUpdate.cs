using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace Havells_Plugin.InventoryRequest
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
                    && context.PrimaryEntityName.ToLower() == hil_inventoryrequest.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    UpdateInventory(entity, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.InventoryRequest.PostUpdate.Execute" + ex.Message);
            }
        }

        public static void UpdateInventory(Entity entity,IOrganizationService service, ITracingService iTrace)
        {
            hil_inventoryrequest enInvReq = entity.ToEntity<hil_inventoryrequest>();
            if (enInvReq.statuscode.Value == 910590000)//submitted91,05,90,000
            {
                hil_inventoryrequest InvReq =(hil_inventoryrequest)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));

                Guid fsPart = Guid.Empty;
                Guid fsOwnerAccount = Guid.Empty;
                Guid fsOwner = Guid.Empty;
                OptionSetValue opInventoryType = new OptionSetValue(2); //out warranty
                Int32 iGoodQty = 0;
                Guid fsInvReqId= Guid.Empty;
                if (InvReq.hil_PartCode!= null)
                    fsPart = InvReq.hil_PartCode.Id;
                if (InvReq.hil_Account != null)
                    fsOwnerAccount = InvReq.hil_Account.Id;
                if (InvReq.OwnerId != null)
                    fsOwner = InvReq.OwnerId.Id;
                if (InvReq.hil_Quantity!= null)
                    iGoodQty = InvReq.hil_Quantity.Value;
                fsInvReqId = InvReq.Id;
                HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, Guid.Empty,Guid.Empty, fsInvReqId,Guid.Empty, Guid.Empty, service, iTrace, iGoodQty);

            }
        }

    }
}
