using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;

namespace Havells_Plugin.PerformaInvoice
{
    public class PreUpdate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_claimheader" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity PreImage = ((Entity)context.PreEntityImages["image"]);

                    if (entity.Attributes.Contains("hil_performastatus"))
                    {
                        OptionSetValue _performaStatus = entity.GetAttributeValue<OptionSetValue>("hil_performastatus");
                        if (_performaStatus.Value == 3) //Approved
                        {
                            entity["hil_approvedby"] = new EntityReference("systemuser", context.UserId);
                            entity["hil_approvedon"] = DateTime.Now.AddMinutes(330);
                        }
                    }
                    if (PreImage.Attributes.Contains("hil_performastatus"))
                    {
                        OptionSetValue _performaStatus = PreImage.GetAttributeValue<OptionSetValue>("hil_performastatus");
                        if (_performaStatus.Value == 4) //Posted
                        {
                            throw new InvalidPluginExecutionException("Access Denied!!! Performa Invoice is already Posted.");
                        }
                    }
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }

        }
    }
}
