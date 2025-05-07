using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.IPGeo
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
                    && context.PrimaryEntityName.ToLower() == hil_politicalmapping.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE")
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    SetName(entity, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.IPGeo.PostCreate.PostUpdate.Execute" + ex.Message);
            }
        }
        public static void SetName(Entity entity,IOrganizationService service)
        {
            try
            {

                tracingService.Trace("2");
                hil_politicalmapping IPGeo = entity.ToEntity<hil_politicalmapping>();
                if (IPGeo.hil_pincode!=null || IPGeo.hil_city!=null|| IPGeo.hil_state!=null || IPGeo.hil_district!=null)
                {
                    IPGeo = (hil_politicalmapping)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_pincode", "hil_city", "hil_state", "hil_district"));

                    tracingService.Trace("3");
                    String sName = String.Empty;
                    String sPinCode = String.Empty;
                    String sCity = String.Empty;
                    String sState = String.Empty;
                    String sDistrict = String.Empty;
                    if (IPGeo.hil_pincode != null)
                    { sName += IPGeo.hil_pincode.Name; }
                    if (IPGeo.hil_city != null)
                    { sName += IPGeo.hil_city.Name; }
                    if (IPGeo.hil_state != null)
                    { sName += IPGeo.hil_state.Name; }
                    if (IPGeo.hil_district != null)
                    { sName += IPGeo.hil_district.Name; }

                    tracingService.Trace("4"+sName);
                    hil_politicalmapping upIPGeo = new hil_politicalmapping();
                    upIPGeo.hil_politicalmappingId = IPGeo.hil_politicalmappingId.Value;
                    upIPGeo.hil_name = sName;
                    service.Update(upIPGeo);

                    tracingService.Trace("5");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.IPGeo.PostUpdate.SetName" + ex.Message);
            }

        }

       

    }
}
