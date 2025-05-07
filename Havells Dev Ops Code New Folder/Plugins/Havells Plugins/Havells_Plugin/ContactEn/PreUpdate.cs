using System;
using Microsoft.Xrm.Sdk;

namespace Havells_Plugin.ContactEn
{
    public class PreUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == Contact.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string Mobilephone = entity.GetAttributeValue<string>("mobilephone");
                    if (!string.IsNullOrEmpty(Mobilephone) && !string.IsNullOrWhiteSpace(Mobilephone))
                    {
                        if (Mobilephone.Length != 10)
                        {
                            throw new InvalidPluginExecutionException("MOBILE NUMBER MUST BE 10 DIGIT.");
                        }
                    }

                    Entity preImage = (Entity)context.PreEntityImages["imgpreupdate"];
                    if (preImage.Contains("mobilephone"))
                    {
                        string preImageMobilephone = preImage.GetAttributeValue<string>("mobilephone");
                        if (Mobilephone != preImageMobilephone)
                        {
                            throw new InvalidPluginExecutionException("Access denied !!! Change in Mobile Number is not allowed.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ContactEn.PreUpdate.Execute" + ex.Message);
            }
        }
    }
}
