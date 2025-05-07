using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class CreateDuplicateOrderCheckList : IPlugin
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
            try
            {
                tracingService.Trace("CreateDuplicateOrderCheckList plugin started  ");
                var entityIds = context.InputParameters["OrderCheckListProductId"].ToString();
                var entityName = context.InputParameters["EntityName"].ToString();
                createDuplicate(service, entityIds, entityName);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
        void createDuplicate(IOrganizationService service, string Guids, string entityName)
        {
            string[] guid = Guids.Split(';');
            foreach (string recordid in guid)
            {
                Entity oclPrd = service.Retrieve(entityName, new Guid(recordid), new ColumnSet(true));

                Entity OCLPRD = new Entity(oclPrd.LogicalName);
                if (oclPrd.Contains("hil_department"))
                {
                    OCLPRD["hil_department"] = oclPrd["hil_department"];
                }
                if (oclPrd.Contains("hil_product"))
                {
                    OCLPRD["hil_product"] = oclPrd["hil_product"];
                }
                if (oclPrd.Contains("hil_productdescription"))
                {
                    OCLPRD["hil_productdescription"] = oclPrd["hil_productdescription"];
                }
                if (oclPrd.Contains("hil_lprsmtr"))
                {
                    OCLPRD["hil_lprsmtr"] = oclPrd["hil_lprsmtr"];
                }
                if (oclPrd.Contains("hil_quantity"))
                {
                    OCLPRD["hil_quantity"] = oclPrd["hil_quantity"];
                }
                if (oclPrd.Contains("hil_selectproduct"))
                {
                    OCLPRD["hil_selectproduct"] = oclPrd["hil_selectproduct"];
                }
                if (oclPrd.Contains("hil_hodiscper"))
                {
                    OCLPRD["hil_hodiscper"] = oclPrd["hil_hodiscper"];
                }
                if (oclPrd.Contains("hil_approveddiscount"))
                {
                    OCLPRD["hil_approveddiscount"] = oclPrd["hil_approveddiscount"];
                }
                if (oclPrd.Contains("hil_marginaddedonhoprice"))
                {
                    OCLPRD["hil_marginaddedonhoprice"] = oclPrd["hil_marginaddedonhoprice"];
                }
                if (oclPrd.Contains("hil_hopricespecialconstructions"))
                {
                    OCLPRD["hil_hopricespecialconstructions"] = oclPrd["hil_hopricespecialconstructions"];
                }
                if (oclPrd.Contains("hil_plantcode"))
                {
                    OCLPRD["hil_plantcode"] = oclPrd["hil_plantcode"];
                }
                if (oclPrd.Contains("hil_basicpriceinrsmtr"))
                {
                    OCLPRD["hil_basicpriceinrsmtr"] = oclPrd["hil_basicpriceinrsmtr"];
                }
                if (oclPrd.Contains("hil_totalvalueinrs"))
                {
                    OCLPRD["hil_totalvalueinrs"] = oclPrd["hil_totalvalueinrs"];
                }
                if (oclPrd.Contains("hil_freightcharges"))
                {
                    OCLPRD["hil_freightcharges"] = oclPrd["hil_freightcharges"];
                }
                if (oclPrd.Contains("hil_tolerancelowerlimit"))
                {
                    OCLPRD["hil_tolerancelowerlimit"] = oclPrd["hil_tolerancelowerlimit"];
                }
                if (oclPrd.Contains("hil_toleranceupperlimit"))
                {
                    OCLPRD["hil_toleranceupperlimit"] = oclPrd["hil_toleranceupperlimit"];
                }
                if (oclPrd.Contains("hil_poamount"))
                {
                    OCLPRD["hil_poamount"] = oclPrd["hil_poamount"];
                }
                if (oclPrd.Contains("hil_podiscount"))
                {
                    OCLPRD["hil_podiscount"] = oclPrd["hil_podiscount"];
                }
                if (oclPrd.Contains("hil_porate"))
                {
                    OCLPRD["hil_porate"] = oclPrd["hil_porate"];
                }
                if (oclPrd.Contains("hil_orderchecklistid"))
                {
                    OCLPRD["hil_orderchecklistid"] = oclPrd["hil_orderchecklistid"];
                }
                if (oclPrd.Contains("hil_inspectiontype"))
                {
                    OCLPRD["hil_inspectiontype"] = oclPrd["hil_inspectiontype"];
                }
                if (oclPrd.Contains("hil_tenderproductid"))
                {
                    OCLPRD["hil_tenderproductid"] = oclPrd["hil_tenderproductid"];
                }
                if (oclPrd.Contains("hil_finaloffervalue"))
                {
                    OCLPRD["hil_finaloffervalue"] = oclPrd["hil_finaloffervalue"];
                }
                service.Create(OCLPRD);
            }
        }
    }
}
