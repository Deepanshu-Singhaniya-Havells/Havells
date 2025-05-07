using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WOProduct
{
    public class PostCreate: IPlugin
    {
        #region PluginConfig
       static ITracingService tracingService =null;
        public void Execute(IServiceProvider serviceProvider)
        {
             tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorderproduct"
                    && context.MessageName.ToUpper() == "CREATE" && context.Depth < 2)
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity ent = (Entity)context.InputParameters["Target"];
                    msdyn_workorderproduct WrkPdt = ent.ToEntity<msdyn_workorderproduct>();
                    SetWarrantyType(service, WrkPdt, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            #endregion
        }
        private static void SetWarrantyType(IOrganizationService service, msdyn_workorderproduct WrkPdt, ITracingService Tracing)
        {
            msdyn_workorderproduct iUpdateWrkPdt = new msdyn_workorderproduct();
            iUpdateWrkPdt.Id = WrkPdt.Id;
            OptionSetValue Warranty = new OptionSetValue();
            OptionSetValue isChargable = new OptionSetValue();
            DateTime iDate = new DateTime();
            if(WrkPdt.msdyn_Product != null && WrkPdt.msdyn_CustomerAsset != null && WrkPdt.msdyn_WorkOrderIncident != null)
            {
                OptionSetValue WtyStatus = new OptionSetValue();
                if(WrkPdt.Contains("hil_warrantystatus") && WrkPdt.hil_WarrantyStatus != null)
                {
                    Warranty = WrkPdt.hil_WarrantyStatus;
                }
                if(Warranty.Value != 2)
                {
                    msdyn_customerasset _enAssetWty = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, WrkPdt.msdyn_CustomerAsset.Id, new ColumnSet("hil_warrantystatus", "hil_warrantytilldate"));

                    if (_enAssetWty.Attributes.Contains("hil_warrantytilldate"))
                    {
                        iDate = _enAssetWty.GetAttributeValue<DateTime>("hil_warrantytilldate");
                        if (DateTime.Now.Date <= iDate.Date)
                        {
                            iUpdateWrkPdt.hil_WarrantyStatus = new OptionSetValue(1); //IN Warranty
                            iUpdateWrkPdt["hil_effectiveamount"] = Convert.ToDecimal(0.00);
                            service.Update(iUpdateWrkPdt);
                        }
                        else
                        {
                            EntityReference Part = WrkPdt.msdyn_Product;
                            EntityReference Asset = WrkPdt.msdyn_CustomerAsset;
                            EntityReference Incident = WrkPdt.msdyn_WorkOrderIncident;
                            QueryExpression Query = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_part", ConditionOperator.Equal, Part.Id);
                            Query.Criteria.AddCondition("hil_producttype", ConditionOperator.Equal, 2);
                            Query.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, Asset.Id);
                            Query.Criteria.AddCondition("hil_warrantystartdate", ConditionOperator.NotNull);
                            Query.Criteria.AddCondition("hil_warrantyenddate", ConditionOperator.NotNull);
                            EntityCollection Found = service.RetrieveMultiple(Query);
                            if (Found.Entities.Count > 0)
                            {
                                foreach (hil_unitwarranty UnitWrty in Found.Entities)
                                {
                                    if (DateTime.Now >= UnitWrty.hil_warrantystartdate && DateTime.Now <= UnitWrty.hil_warrantyenddate)
                                    {
                                        iUpdateWrkPdt.hil_WarrantyStatus = new OptionSetValue(1); //IN Warranty
                                        iUpdateWrkPdt["hil_effectiveamount"] = Convert.ToDecimal(0.00);
                                        service.Update(iUpdateWrkPdt);
                                    }
                                    else
                                    {
                                        //Tracing.Trace("12");
                                        iUpdateWrkPdt.hil_WarrantyStatus = new OptionSetValue(2); //OUT Warranty
                                        iUpdateWrkPdt["hil_effectiveamount"] = WrkPdt.hil_PartAmount;
                                        service.Update(iUpdateWrkPdt);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        iUpdateWrkPdt.hil_WarrantyStatus = new OptionSetValue(2);//OUT Warranty
                        iUpdateWrkPdt["hil_effectiveamount"] = WrkPdt.hil_PartAmount;
                        service.Update(iUpdateWrkPdt);
                    }
                }
            }
        }
    }
}