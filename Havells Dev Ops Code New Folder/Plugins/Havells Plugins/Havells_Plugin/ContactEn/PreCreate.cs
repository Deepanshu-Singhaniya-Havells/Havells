using System;
using Microsoft.Xrm.Sdk;

namespace Havells_Plugin.ContactEn
{
    public class PreCreate : IPlugin
    {
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
                    && context.PrimaryEntityName.ToLower() == Contact.EntityLogicalName && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Contact _Cnt = entity.ToEntity<Contact>();
                    if (!entity.Contains("mobilephone") || !entity.Contains("firstname") || !entity.Contains("hil_consumersource")) {
                        throw new InvalidPluginExecutionException("Consumer Name, Mobile Number and Source of Creation is required.");
                    }
                    if(_Cnt.MobilePhone != null)
                    {
                        if(_Cnt.MobilePhone.Length != 10)
                        {
                            throw new InvalidPluginExecutionException("MOBILE NUMBER MUST BE 10 DIGIT");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("MOBILE NUMBER CAN'T BE NULL");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ContactEn.PreCreate.Execute" + ex.Message);
            }
        }
    }
}
