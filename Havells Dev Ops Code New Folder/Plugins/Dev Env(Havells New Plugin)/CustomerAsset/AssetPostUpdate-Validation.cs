using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using HavellsNewPlugin.Helper;

namespace HavellsNewPlugin.CustomerAsset
{
    public class AssetPostUpdate_Validation : IPlugin
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
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    Guid userId = context.UserId;
                    Entity preImageCA = ((Entity)context.PreEntityImages["image"]);
                    if (entity.Contains("hil_customer"))
                    {
                        EntityReference customer = entity.GetAttributeValue<EntityReference>("hil_customer");
                        if (preImageCA.Contains("hil_customer"))
                        {
                            EntityReference customerImg = preImageCA.GetAttributeValue<EntityReference>("hil_customer");
                            Guid user = context.UserId;
                            if (customer.Id != customerImg.Id)
                            {
                                string position = HelperClass.getUserPosition(user, service, tracingService);
                                if (position == "DSE" || position == "Franchise" || position == "Franchise Technician" || position == null)
                                    throw new InvalidPluginExecutionException("You are not authorized to update Customer, please contact to your Branch Office.");
                            }
                        }
                    }
                    if (entity.Contains("msdyn_name"))
                    {
                        if (HelperClass.getUserSecurityRole(userId, service, "Customer Asset Delink", tracingService))
                        {
                            string srNo = entity.GetAttributeValue<string>("msdyn_name");//a_svc
                            if (srNo.Contains("_SVC"))
                            {
                                srNo = srNo.Replace("_SVC", "");//abc123
                            }


                            string srNoPreImg = preImageCA.GetAttributeValue<string>("msdyn_name").Replace("_SVC", "");//a_svc

                            if (srNo != srNoPreImg)
                            {
                                throw new InvalidPluginExecutionException("You are not authorized to update  Serial Number, please contact to your Branch Office.");
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("You are not authorized to update  Serial Number, please contact to your Branch Office.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
            #endregion
        }

    }
}
