using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Havells_Plugin;

namespace Plugins.Job_Services
{
    public class PreCreate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderservice.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity ent = (Entity)context.InputParameters["Target"];
                    msdyn_workorderservice WrkService = ent.ToEntity<msdyn_workorderservice>();
                    // Added by Kuldeep Khare 31/Dec/2019
                    // Check : Work Order/Work Order Incident should be mandatory
                    if (WrkService.msdyn_WorkOrderIncident.Id == Guid.Empty) {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - Work Order Incident is mandatory - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                    if (WrkService.msdyn_WorkOrder.Id == Guid.Empty)
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - Work Order is mandatory - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }

                    msdyn_workorderincident iIncident = (msdyn_workorderincident)service.Retrieve(msdyn_workorderincident.EntityLogicalName, WrkService.msdyn_WorkOrderIncident.Id, new ColumnSet("new_warrantyenddate", "msdyn_workorder", "msdyn_customerasset"));

                    GetActionWarranty(service, iIncident, WrkService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("JobService.PreCreate " + ex.Message);
            }
            #endregion
        }
        #region Get Warranty Status
        public static void GetActionWarranty(IOrganizationService service, msdyn_workorderincident iInc, msdyn_workorderservice iWoServ)
        {
            if (iInc.msdyn_WorkOrder != null)
            {
                msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, iInc.msdyn_WorkOrder.Id, new ColumnSet("createdon"));
                DateTime iWEndDate = new DateTime();
                DateTime dtCompare = new DateTime();

                if (iInc.Contains("new_warrantyenddate") && iInc.Attributes.Contains("new_warrantyenddate"))
                {
                    iWEndDate = iInc.GetAttributeValue<DateTime>("new_warrantyenddate"); //(DateTime)iInc["new_warrantyenddate"];
                }
                else
                {
                    msdyn_customerasset Customerasset = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, iInc.msdyn_CustomerAsset.Id,
                        new ColumnSet(new string[] { "hil_warrantytilldate" }));
                    if (Customerasset.Contains("hil_warrantytilldate"))
                    {
                        iWEndDate = Customerasset.GetAttributeValue<DateTime>("hil_warrantytilldate");
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("Warranty till data on custome assest is blank");
                    }
                }
                dtCompare = (DateTime)enJob.CreatedOn;
                if (dtCompare <= iWEndDate)
                {
                    if (iWoServ.msdyn_Service != null)
                    {
                        Product _JobService = (Product)service.Retrieve(Product.EntityLogicalName, iWoServ.msdyn_Service.Id, new ColumnSet("hil_amount"));
                        if (_JobService.hil_Amount != null)
                        {
                            iWoServ.hil_WarrantyStatus = new OptionSetValue(1); //IN Warranty
                            iWoServ["hil_charge"] = _JobService.hil_Amount.Value;
                            iWoServ["hil_effectivecharge"] = Convert.ToDecimal(0);
                        }
                    }
                }
                else
                {
                    if (iWoServ.msdyn_Service != null)
                    {
                        Product _JobService = (Product)service.Retrieve(Product.EntityLogicalName, iWoServ.msdyn_Service.Id, new ColumnSet("hil_amount"));
                        if (_JobService.hil_Amount != null)
                        {
                            iWoServ.hil_WarrantyStatus = new OptionSetValue(2); //OUT Warranty
                            iWoServ["hil_charge"] = _JobService.hil_Amount.Value;
                            iWoServ["hil_effectivecharge"] = Convert.ToDecimal(_JobService.hil_Amount.Value);
                        }
                    }
                }
            }
        }

        public static void SetActionWarrantyStatus(IOrganizationService service, msdyn_workorderincident iInc, msdyn_workorderservice iWoServ)
        {
            if (iInc.msdyn_WorkOrder != null)
            {
                msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, iInc.msdyn_WorkOrder.Id, new ColumnSet("createdon"));
                Decimal _partAmount = 0;

                ApplicableWarrantyDTO _retObj = WarrantyEngine.ApplicationOfCustomerAssetWarranty(service, new ApplicableWarrantyDTO()
                {
                    jobCreatedOn = Convert.ToDateTime(enJob.CreatedOn),
                    customerAssetId = iInc.msdyn_CustomerAsset.Id,
                    replacesPartId = Guid.Empty
                });
                Product _JobService = (Product)service.Retrieve(Product.EntityLogicalName, iWoServ.msdyn_Service.Id, new ColumnSet("hil_amount"));
                if (_JobService.hil_Amount != null) {
                    _partAmount = _JobService.hil_Amount.Value;
                }
                iWoServ.hil_WarrantyStatus = _retObj.LaborWarrantyStatus;
                iWoServ["hil_charge"] = _partAmount;
                iWoServ["hil_effectivecharge"] = _retObj.LaborWarrantyStatus.Value == 1 ? Convert.ToDecimal(0) : _partAmount;
            }
        }
        #endregion
    }
}