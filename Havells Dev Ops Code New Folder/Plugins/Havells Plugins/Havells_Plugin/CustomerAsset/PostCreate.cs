using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk.Query;
namespace Havells_Plugin.CustomerAsset
{
    public class PostCreate : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
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
                    && context.PrimaryEntityName.ToLower() == "msdyn_customerasset"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Attributes.Contains("hil_createwarranty"))
                    {
                        bool IsCreateWrty = (bool)entity["hil_createwarranty"];
                        if (IsCreateWrty == true)
                        {
                            //HelperWarrantyModule.Init_Warranty(entity.Id, service);
                            WarrantyEngine.ExecuteWarrantyEngine(service, entity.Id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PostCreate.Execute" + ex.Message);
            }
            #endregion
        }
        //public static void CreateWorkOrder(IOrganizationService service, Guid AsstId)
        //{
        //    try
        //    { 
        //    Guid CallSubtype = Helper.GetGuidbyName(hil_callsubtype.EntityLogicalName, "hil_name", "Product Registration", service);
        //    Guid CallType = Helper.GetGuidbyName(hil_calltype.EntityLogicalName, "hil_name", "Service", service);
        //    Guid Observation = Helper.GetGuidbyName(hil_observation.EntityLogicalName, "hil_name", "Product Registration", service);
        //    Guid Nature = Helper.GetGuidbyName(hil_natureofcomplaint.EntityLogicalName, "hil_name", "Product Registration", service);
        //    Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
        //    Guid IncidentType = Helper.GetGuidbyName(msdyn_incidenttype.EntityLogicalName, "msdyn_name", "Product Registration", service);
        //    //Contact 
        //    msdyn_customerasset CustAsst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, AsstId, new ColumnSet(true));
        //    msdyn_workorder Wo = new msdyn_workorder();
        //    Wo.msdyn_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, CustAsst.Id);
        //    if (CustAsst.hil_Customer != null)
        //        Wo.hil_CustomerRef = CustAsst.hil_Customer;
        //    if (CustAsst.msdyn_Account != null)
        //    {
        //        Wo.msdyn_ServiceAccount = CustAsst.msdyn_Account;
        //        Wo.msdyn_BillingAccount = CustAsst.msdyn_Account;
        //    }
        //    if (IncidentType != Guid.Empty)
        //    {
        //        Wo.msdyn_PrimaryIncidentType = new EntityReference(msdyn_incidenttype.EntityLogicalName, IncidentType);
        //    }
        //    if (CustAsst.hil_ProductCategory != null)
        //        Wo.hil_Productcategory = CustAsst.hil_ProductCategory;
        //    if (CallSubtype != Guid.Empty)
        //    {
        //        Wo.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, CallSubtype);
        //    }
        //    if (CallType != Guid.Empty)
        //    {
        //        Wo.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, CallType);
        //    }
        //    if (Observation != Guid.Empty)
        //    {
        //        Wo.hil_observation = new EntityReference(hil_observation.EntityLogicalName, Observation);
        //    }
        //    if (Nature != Guid.Empty)
        //    {
        //        Wo.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, Nature);
        //    }
        //    if (PriceList != Guid.Empty)
        //    {
        //        Wo.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
        //    }
        //    service.Create(Wo);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PostCreate.CreateWorkOrder" + ex.Message);
        //    }
        //}
    }
}