using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WOProduct
{
    public class PostUpdate_JobProduct_ReplacedPart_Warranty : IPlugin
    {
        #region PluginConfig
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderproduct.EntityLogicalName.ToLower() && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity ent = (Entity)context.InputParameters["Target"];
                    msdyn_workorderproduct iUpdate = ent.ToEntity<msdyn_workorderproduct>();
                    if(ent.Contains("hil_replacedpart") && ent.Attributes.Contains("hil_replacedpart"))
                    {
                        tracingService.Trace("1 - Execute : " + DateTime.Now);
                        msdyn_workorderproduct preEntityJobProduct = context.PreEntityImages["JobProductPreEntity"].ToEntity<msdyn_workorderproduct>();
                        GetWarranty(service, preEntityJobProduct, iUpdate, tracingService);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.WorkOrderProduct.ReplacedPart_Warranty" + ex.Message);
            }
            #endregion
        }
        #endregion
        public static void GetWarranty(IOrganizationService service, msdyn_workorderproduct WrkPdt, msdyn_workorderproduct iUpdate,ITracingService Tracing)
        {
            try
            {
                Tracing.Trace("1 - Get Warranty : " + DateTime.Now);
                hil_unitwarranty UnitWrty = new hil_unitwarranty();
                DateTime iDate = new DateTime();
                if (WrkPdt.msdyn_Product != null && WrkPdt.msdyn_CustomerAsset != null && WrkPdt.msdyn_WorkOrderIncident != null)
                {
                    OptionSetValue WtyStatus = new OptionSetValue();
                    if(WrkPdt.Contains("hil_warrantystatus") && WrkPdt.hil_WarrantyStatus != null)
                    {
                        WtyStatus = WrkPdt.hil_WarrantyStatus;
                    }
                    if(WtyStatus.Value != 2)
                    {
                        msdyn_customerasset _enAssetWty = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, WrkPdt.msdyn_CustomerAsset.Id, new ColumnSet("hil_warrantystatus", "hil_warrantytilldate"));
                        if (_enAssetWty.Attributes.Contains("hil_warrantytilldate"))
                        {
                            Tracing.Trace("2 - Get Warranty : " + DateTime.Now);
                            iDate = _enAssetWty.GetAttributeValue<DateTime>("hil_warrantytilldate");
                            if (DateTime.Now <= iDate)
                            {
                                Tracing.Trace("3 - Get Warranty : " + DateTime.Now);
                                iUpdate.hil_WarrantyStatus = new OptionSetValue(1);
                                iUpdate["hil_effectiveamount"] = Convert.ToDecimal(0.00);
                            }
                            else
                            {
                                Tracing.Trace("4 - Get Warranty : " + DateTime.Now);
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
                                Query.AddOrder("hil_warrantyenddate", OrderType.Descending);
                                EntityCollection Found = service.RetrieveMultiple(Query);
                                if (Found.Entities.Count > 0)
                                {
                                    UnitWrty = Found.Entities[0].ToEntity<hil_unitwarranty>();
                                    if (DateTime.Now >= UnitWrty.hil_warrantystartdate && DateTime.Now <= UnitWrty.hil_warrantyenddate)
                                    {
                                        Tracing.Trace("5 - Get Warranty : " + DateTime.Now);
                                        iUpdate.hil_WarrantyStatus = new OptionSetValue(1);
                                        iUpdate["hil_effectiveamount"] = Convert.ToDecimal(0.00);
                                    }
                                    else
                                    {
                                        Tracing.Trace("6 - Get Warranty : " + DateTime.Now);
                                        iUpdate.hil_WarrantyStatus = new OptionSetValue(2);
                                        iUpdate["hil_effectiveamount"] = WrkPdt.hil_PartAmount;
                                    }
                                }
                                else
                                {
                                    Tracing.Trace("6 - Get Warranty : " + DateTime.Now);
                                    iUpdate.hil_WarrantyStatus = new OptionSetValue(2);
                                    iUpdate["hil_effectiveamount"] = WrkPdt.hil_PartAmount;
                                }
                            }
                        }
                        else
                        {
                            Tracing.Trace("7 - Get Warranty : " + DateTime.Now);
                            iUpdate.hil_WarrantyStatus = new OptionSetValue(2);
                            iUpdate["hil_effectiveamount"] = WrkPdt.hil_PartAmount;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("WorkOrderProduct.ReplacedPart_Warranty___ERROR : " + ex.Message.ToUpper());
            }
        }
    }
}