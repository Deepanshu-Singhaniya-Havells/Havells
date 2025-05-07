using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class PostUpdateTenderProduct : IPlugin
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
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_tenderproduct")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity entObj = service.Retrieve("hil_tenderproduct", entity.Id, new ColumnSet("hil_lprsmtr", "hil_hopricespecialconstructions", "hil_hodiscper"));
                    //if (entObj.Contains("hil_hopricespecialconstructions"))
                    //{
                    //    Entity ent = service.Retrieve("hil_tenderproduct", entity.Id, new ColumnSet("hil_lprsmtr"));
                    //    if (ent!=null && ent.Contains("hil_lprsmtr"))
                    //    {
                    //        Decimal _lpMtr = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                    //        Decimal _hoPrice = entity.GetAttributeValue<Money>("hil_hopricespecialconstructions").Value;
                    //        Decimal hodiscount = ((1 - _hoPrice / _lpMtr) * 100);
                    //        Entity entUpdate = new Entity("hil_tenderproduct", entity.Id);
                    //        entUpdate["hil_approveddiscount"] = hodiscount;
                    //        entUpdate["hil_hodiscper"] = hodiscount;
                    //        service.Update(entUpdate);
                    //    }
                    //}
                    if (entObj.Contains("hil_hodiscper"))
                    {
                        Decimal _hoDiscPer = Math.Round(entity.GetAttributeValue<Decimal>("hil_hodiscper"),2);
                        Entity entUpdate = new Entity("hil_tenderproduct", entity.Id);
                        entUpdate["hil_approveddiscount"] = _hoDiscPer;

                        Entity ent = service.Retrieve("hil_tenderproduct", entity.Id, new ColumnSet("hil_lprsmtr", "hil_hopricespecialconstructions"));
                        if (ent != null && ent.Contains("hil_lprsmtr"))
                        {
                            Decimal _lpMtr = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                            Decimal hoPrice = Math.Round((_lpMtr * (1 - _hoDiscPer / 100)), 2);
                            entUpdate["hil_hopricespecialconstructions"] = new Money(hoPrice);
                        }
                        service.Update(entUpdate);
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.PostUpdateTenderProduct.Execute Error " + ex.Message);
            }
        }
    }
}
