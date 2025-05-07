using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.SAWActivityApproval
{
    public class PreCreate : IPlugin
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
                    && context.PrimaryEntityName.ToLower() == "hil_sawactivityapproval"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");

                    if (entity.Attributes.Contains("hil_sawactivity"))
                    {
                        tracingService.Trace("2");
                        var sawActivity = ((EntityReference)(entity.Attributes["hil_sawactivity"]));
                        var entSAWActivity = service.Retrieve(sawActivity.LogicalName, sawActivity.Id, new ColumnSet("hil_name"));
                        string sawActivityNo = entSAWActivity["hil_name"].ToString();
                        string levelNo = string.Empty;

                        if (entity.Attributes.Contains("hil_level"))
                        {
                            OptionSetValue _optValue = ((OptionSetValue)(entity.Attributes["hil_level"]));
                            levelNo = _optValue.Value == 1 ? "L1" : _optValue.Value == 2 ? "L2" : "L3";
                        }
                        tracingService.Trace("");
                        entity["hil_name"] = sawActivityNo + levelNo;
                        tracingService.Trace("3");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SAWActivity.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }
    }
}
