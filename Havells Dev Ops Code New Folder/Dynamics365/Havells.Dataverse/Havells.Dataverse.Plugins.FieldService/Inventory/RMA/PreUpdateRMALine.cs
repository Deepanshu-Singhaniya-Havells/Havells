using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Services;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdateRMALine : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventoryrmaline" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    Entity _entityLine = service.Retrieve("hil_inventoryrmaline", entity.Id, new ColumnSet("hil_rma"));

                    EntityReference _rmaHeader = _entityLine.GetAttributeValue<EntityReference>("hil_rma");
                    Entity entHeader = service.Retrieve(_rmaHeader.LogicalName, _rmaHeader.Id, new ColumnSet("hil_inspectionnumber"));
                    if (entHeader.Contains("hil_inspectionnumber"))
                    {
                        string _inspectionNum = entHeader.GetAttributeValue<string>("hil_inspectionnumber");
                        if (!string.IsNullOrWhiteSpace(_inspectionNum))
                        {
                            throw new InvalidPluginExecutionException("You can't Update RMA Line item as Inspection is already submitted!!");
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
}
