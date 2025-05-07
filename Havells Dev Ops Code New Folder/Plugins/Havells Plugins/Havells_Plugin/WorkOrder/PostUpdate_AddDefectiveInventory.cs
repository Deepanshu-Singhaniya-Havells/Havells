using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WorkOrder
{
    public class PostUpdate_AddDefectiveInventory : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Plugin Configuration
            ITracingService tracingService1 = null;
            tracingService1 = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();

                    tracingService1.Trace($@"Step 1: Start Increasing Defective Inv . {DateTime.Now.ToString()}");

                    IncreaseDefectiveInv(service, entity, tracingService1);

                    tracingService1.Trace($@"Step 2: Completed Increasing Defective Inv . {DateTime.Now.ToString()}");
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate_AddDefectiveInventory.Execute: " + ex.Message);
            }
            #endregion
        }
        #region Increase Defective Inventory
        public static void IncreaseDefectiveInv(IOrganizationService service, Entity entity, ITracingService tracingService1)
        {
            try
            {
                tracingService1.Trace($@"Step 1.1: Start Increasing Defective Inv . {DateTime.Now.ToString()}");
                msdyn_workorder Job = entity.ToEntity<msdyn_workorder>();
                //if Job complete==yes
                if (Job.hil_CalculateCharges != null && Job.hil_CalculateCharges.Value == true)
                {
                    //iTrace.Trace("1.2");

                    tracingService1.Trace($@"Step 1.2: Deleting all Unsed Parts . {DateTime.Now.ToString()}");

                    //DeleteAllUnusedPartsandService(service, entity.Id);

                    //iTrace.Trace("1.3");
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {

                        var obj = from _woProduct in orgContext.CreateQuery<msdyn_workorderproduct>()
                                join _WorkOrder in orgContext.CreateQuery<msdyn_workorder>() on _woProduct.msdyn_WorkOrder.Id equals _WorkOrder.Id
                                //  join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.msdyn_ServiceAccount.Id equals _Account.Id
                                join _Account in orgContext.CreateQuery<Account>() on _WorkOrder.hil_OwnerAccount.Id equals _Account.Id
                                join _Owner in orgContext.CreateQuery<SystemUser>() on _Account.OwnerId.Id equals _Owner.Id
                                where _woProduct.msdyn_WorkOrder.Id == entity.Id
                                && _woProduct.hil_replacedpart != null
                                select new
                                {
                                    _woProduct.msdyn_workorderproductId,
                                    _woProduct.hil_replacedpart,
                                    _woProduct.msdyn_Quantity,
                                    _woProduct.hil_AvailabilityStatus,
                                    _woProduct.hil_WarrantyStatus,
                                    _woProduct.statecode,
                                    _Account.AccountId,
                                    _Owner.SystemUserId,
                                    _Account.Name,
                                    _Owner.FullName
                                };
                        tracingService1.Trace($@"Step 1.3: Looping Through the fetched list of Spare Parts . {DateTime.Now.ToString()}");
                        foreach (var iobj in obj)
                        {
                            //iTrace.Trace("1.3");
                            //tracingService.Trace("UI5 account name " + iobj.Name + " Owner " + iobj.FullName);
                            Guid fsPart = Guid.Empty;
                            Guid fsOwnerAccount = Guid.Empty;
                            Guid fsOwner = Guid.Empty;
                            OptionSetValue opInventoryType = null; //to do need value
                            Int32 iQuantityRequired = 0;
                            Guid fsWOProductId = iobj.msdyn_workorderproductId.Value;
                            if (iobj.statecode.Equals(msdyn_workorderproductState.Active))
                            {
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
                                    opInventoryType = iobj.hil_WarrantyStatus;

                                //Update defective inventory of current available item
                                {
                                    tracingService1.Trace($@"Step 1.4: Adding Spare Part Inventory in Inventory Journal . {DateTime.Now.ToString()}");
                                    //HelperInvJournal.CreateInvJournal(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, -iQuantityRequired, iQuantityRequired);
                                    //HelperInvJournal.CreateInvJournalDefective(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, 0, iQuantityRequired);
                                    HelperInvJournal.CreateInvJournalDefective(fsPart, fsOwnerAccount, fsOwner, opInventoryType, fsWOProductId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, service, tracingService1, iQuantityRequired);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.IncreaseDefectiveInv: " + ex.Message);
            }

        }
        #endregion
        #region DeleteAllUnusedPartsandService
        public static void DeleteAllUnusedPartsandService(IOrganizationService service, Guid JobId)
        {
            try
            {
                String fetchPart = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorderproduct'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_product' />
                    <attribute name='msdyn_linestatus' />
                    <attribute name='msdyn_description' />
                    <attribute name='msdyn_workorderproductid' />
                    <order attribute='msdyn_product' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_replacedpart' operator='null' />
                      <condition attribute='msdyn_workorder' operator='eq'  value='{0}' />
                    </filter>
                  </entity>
                </fetch>";
                fetchPart = String.Format(fetchPart, JobId);
                EntityCollection enCollProduct = service.RetrieveMultiple(new FetchExpression(fetchPart));
                foreach (Entity en in enCollProduct.Entities)
                {
                    //service.Delete(en.LogicalName, en.Id);
                }

                String fetchService = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorderservice'>
                        <attribute name='createdon' />
                        <attribute name='msdyn_workorder' />
                        <attribute name='msdyn_name' />
                        <attribute name='msdyn_workorderserviceid' />
                        <order attribute='msdyn_name' descending='false' />
                        <filter type='and'>
                          <condition attribute='msdyn_linestatus' operator='ne' value='690970001' />
                          <condition attribute='msdyn_workorder' operator='eq'   value='{0}' />
                        </filter>
                      </entity>
                    </fetch>";
                fetchService = String.Format(fetchService, JobId);
                EntityCollection enCollService = service.RetrieveMultiple(new FetchExpression(fetchService));
                foreach (Entity en in enCollService.Entities)
                {
                    //service.Delete(en.LogicalName, en.Id);
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(" Error in deleting un-unsed parts and services" + ex.Message);
            }

        }
        #endregion
    }
}