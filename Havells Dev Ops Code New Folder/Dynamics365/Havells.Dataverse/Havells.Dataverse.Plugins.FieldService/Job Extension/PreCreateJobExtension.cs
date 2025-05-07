using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Diagnostics.Eventing.Reader;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreateJobExtension : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_jobsextension" && (context.MessageName.ToUpper() == "CREATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                ProcessRequest(entity, _tracingService, service);
            }
        }

        private void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                if (entity.Contains("hil_jobs"))
                {
                    Entity _jobHeader = service.Retrieve("msdyn_workorder", entity.GetAttributeValue<EntityReference>("hil_jobs").Id, new ColumnSet("hil_preferreddate", "hil_preferredtime"));
                    if (_jobHeader.Contains("hil_preferreddate"))
                    {
                        int _hour = 19; //8:00 PM
                        DateTime _preferredDate = _jobHeader.GetAttributeValue<DateTime>("hil_preferreddate").AddMinutes(330);
                        OptionSetValue _preferredTime = _jobHeader.Contains("hil_preferredtime") ? _jobHeader.GetAttributeValue<OptionSetValue>("hil_preferredtime") : null;
                        _hour = _preferredTime.Value == 1 ? 11 : _preferredTime.Value == 2 ? 15 : 19;

                        entity["hil_preferreddatetime"] = new DateTime(_preferredDate.Year, _preferredDate.Month, _preferredDate.Day, _hour, 59, 59);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
