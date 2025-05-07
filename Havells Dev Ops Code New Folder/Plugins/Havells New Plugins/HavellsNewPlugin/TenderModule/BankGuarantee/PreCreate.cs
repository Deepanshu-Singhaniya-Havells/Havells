using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.TenderModule.BankGuarantee
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
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    string sufix = "EMD";
                    Entity entity = (Entity)context.InputParameters["Target"];

                    int bgType = entity.GetAttributeValue<OptionSetValue>("hil_purpose").Value;
                    string autoNumber = entity.GetAttributeValue<string>("hil_name");
                    if (bgType == 910590000)
                    {
                        sufix = "SDBG";
                    }
                    else if (bgType == 910590001)
                    {
                        sufix = "PBG";
                    }
                    else if (bgType == 910590002)
                    {
                        sufix = "ABG";
                    }
                    else if (bgType == 910590003)
                    {
                        sufix = "EMD";
                    }
                    else if (bgType == 910590004)
                    {
                        sufix = "PMD";
                    }
                    autoNumber = autoNumber.Replace("EMD", sufix);
                    entity["hil_name"] = autoNumber;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.BankGuarantee Error " + ex.Message);
            }
        }
    }
}
