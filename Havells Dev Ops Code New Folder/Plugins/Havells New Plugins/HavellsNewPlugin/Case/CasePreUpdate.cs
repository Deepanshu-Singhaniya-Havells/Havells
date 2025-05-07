using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    internal class CasePreUpdate
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
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
              && context.PrimaryEntityName.ToLower() == "incident")
            {
                Entity entity = (Entity)context.InputParameters["Target"];

                if (entity.Contains("statuscode"))
                {
                    OptionSetValue statusCode = entity.GetAttributeValue<OptionSetValue>("statuscode");

                    if(statusCode.Value == 6)// Cancellation 
                    {
                        throw new InvalidPluginExecutionException("Cancellation of the case is not allowed");
                    }
                }
            }
        }
    }
}

