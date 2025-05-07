using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Address
{
    public class AddressMDMAsync : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                tracingService.Trace("Depth is " + context.Depth);
                if (context.Depth == 1)
                    if (context.InputParameters.Contains("Target")
                       && context.InputParameters["Target"] is Entity
                       && context.PrimaryEntityName.ToLower() == "hil_address"
                       )
                    {
                        tracingService.Trace("1");
                        if (context.MessageName.ToUpper() == "CREATE")
                        {
                            tracingService.Trace("Create ");
                            Entity _entity = (Entity)context.InputParameters["Target"];
                            Entity contactEntity = service.Retrieve("hil_address", _entity.Id, new ColumnSet(true));
                            ContactEn.Common.CallAPIInsertUpdateAddress(service, contactEntity, tracingService);
                            tracingService.Trace("1");
                        }
                        else if (context.MessageName.ToUpper() == "UPDATE")
                        {
                            tracingService.Trace("Update ");
                            Entity _entityTar = (Entity)context.InputParameters["Target"];
                            Entity contactEntity = service.Retrieve("hil_address", _entityTar.Id, new ColumnSet(true));
                            ContactEn.Common.CallAPIInsertUpdateAddress(service, contactEntity, tracingService);
                            tracingService.Trace("1");
                        }
                    }
            }
            catch (Exception ex)
            {
                // throw new InvalidPluginExecutionException("Error : " + ex.Message);
                Entity _address = new Entity(((Entity)context.InputParameters["Target"]).LogicalName);
                _address["hil_errormsg"] = ex.Message;
                _address.Id = ((Entity)context.InputParameters["Target"]).Id;
                service.Update(_address);
                //throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}
