using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.WorkOrder
{
    public class KKGCodeUpdate :IPlugin 
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity preImageWO = ((Entity)context.PreEntityImages["image"]);
                    if (preImageWO.Contains("hil_kkgotp"))
                    {
                        string _oldKkg = preImageWO.GetAttributeValue<string>("hil_kkgotp");
                        string _newKkg = entity.GetAttributeValue<string>("hil_kkgotp");
                        if (_newKkg != _oldKkg)
                        {
                            entity["hil_kkgotp"] = _oldKkg;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}