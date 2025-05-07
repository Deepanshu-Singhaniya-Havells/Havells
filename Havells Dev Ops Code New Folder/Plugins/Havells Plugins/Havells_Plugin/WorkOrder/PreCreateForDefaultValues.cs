using System;
using Havells_Plugin;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WorkOrder
{
    public class PreCreateForDefaultValues : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    SetDefaultValues(entity, service);

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.SetAutoNumber: " + ex.Message);
            }
            #endregion
        }
        public static void SetDefaultValues(Entity _wOd, IOrganizationService service)
        {
            try
            {
                EntityReference CallSubType = new EntityReference();
                Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
                Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                if (_wOd.Attributes.Contains("hil_callsubtype"))
                {
                    CallSubType = (EntityReference)_wOd["hil_callsubtype"];
                }
                EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
                hil_callsubtype ClSb = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, CallSubType.Id, new ColumnSet("hil_calltype"));
                _wOd["msdyn_serviceaccount"] = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                _wOd["msdyn_billingaccount"] = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                if (PriceList != Guid.Empty)
                {
                    _wOd["msdyn_pricelist"] = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
                }
                if (ClSb.hil_CallType != null)
                {
                    CallType = ClSb.hil_CallType;
                    _wOd["hil_calltype"] = CallType;
                }
                if (_wOd.Attributes.Contains("hil_address"))
                {
                    EntityReference erAddress = _wOd.GetAttributeValue<EntityReference>("hil_address");// (EntityReference)_wOd["hil_address"];
                    hil_address enAddress = (hil_address)service.Retrieve(hil_address.EntityLogicalName, erAddress.Id, new ColumnSet("hil_pincode"));
                    if(enAddress.hil_PinCode != null)
                    {
                        QueryByAttribute Qry = new QueryByAttribute();
                        Qry.EntityName = hil_businessmapping.EntityLogicalName;
                        ColumnSet Col = new ColumnSet("hil_region", "hil_branch", "hil_salesoffice");
                        Qry.ColumnSet = Col;
                        Qry.AddAttributeValue("hil_pincode", enAddress.hil_PinCode.Id);
                        EntityCollection Found = service.RetrieveMultiple(Qry);
                        if (Found.Entities.Count > 0)
                        {
                            foreach (hil_businessmapping Bus in Found.Entities)
                            {
                                if (Bus.hil_branch != null)
                                {
                                    _wOd["hil_branch"] = Bus.hil_branch;
                                }
                                if (Bus.hil_region != null)
                                {
                                    _wOd["hil_region"] = Bus.hil_region;
                                }
                                if(Bus.hil_salesoffice != null)
                                {
                                    _wOd["hil_salesoffice"] = enAddress.hil_SalesOffice;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.WorkOrder_Pre_Create.SetAutoNumber: " + ex.Message);
            }
        }
    }
}