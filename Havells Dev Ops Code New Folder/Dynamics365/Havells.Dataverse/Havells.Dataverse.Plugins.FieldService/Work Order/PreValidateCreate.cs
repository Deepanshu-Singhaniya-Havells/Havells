using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Work_Order
{
    public class PreValidateCreate : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    ProcessRequest(entity, _tracingService, service);
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
        public void ProcessRequest(Entity entity, ITracingService _tracingService,IOrganizationService service)
        {
            if (entity.Contains("hil_preferreddate") && entity.Contains("hil_sourceofjob"))
            {
                OptionSetValue _sourceofjob = entity.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                if (!entity.Contains("hil_preferredtime"))
                {
                    throw new InvalidPluginExecutionException("Preferred Day Breakup (Morning|AfterNoon|Evening) is required.");
                }
                else
                {
                    if (_sourceofjob.Value == 24) {
                        return;
                    }
                    DateTime _jobCreationDT = DateTime.Now.AddMinutes(330);
                    int _hour = 0;
                    DateTime _preferredDate = entity.GetAttributeValue<DateTime>("hil_preferreddate").Date;
                    OptionSetValue _preferredTime = entity.GetAttributeValue<OptionSetValue>("hil_preferredtime");
                    _hour = _preferredTime.Value == 1 ? 12 : _preferredTime.Value == 2 ? 16 : 20;
                    DateTime _preferredVisitDT = _preferredDate.AddHours(_hour);

                    TimeSpan diff = _preferredVisitDT - _jobCreationDT;
                    double hours = diff.TotalHours;

                    if (hours < 8)
                    {
                        throw new InvalidPluginExecutionException($"Preferred Datetime must be on and after {_jobCreationDT.AddDays(1).ToString()}");
                    }
                }
            }
        }
    }
}
