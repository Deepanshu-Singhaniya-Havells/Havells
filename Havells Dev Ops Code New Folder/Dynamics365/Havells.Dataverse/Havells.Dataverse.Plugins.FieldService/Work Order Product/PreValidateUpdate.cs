using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.WorkOrderProduct
{
    public class PreValidateUpdate : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "msdyn_workorderproduct" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("msdyn_product"))
                        throw new InvalidPluginExecutionException("Access denied! You can't cnange Default Defective Spare Part.");

                    if (entity.Contains("hil_replacedpart"))
                    {
                        EntityReference _jobProductEntRef = entity.GetAttributeValue<EntityReference>("hil_replacedpart");
                        if (_jobProductEntRef != null && _jobProductEntRef.Id != Guid.Empty)
                        {
                            Entity _jobProductEnt = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("msdyn_workorderincident"));
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_workorderproduct'>
                                <attribute name='msdyn_workorderproductid' />
                                <filter type='and'>
                                  <condition attribute='msdyn_workorderincident' operator='eq' value='{_jobProductEnt.GetAttributeValue<EntityReference>("msdyn_workorderincident").Id}' />
                                  <condition attribute='hil_replacedpart' operator='eq' value='{_jobProductEntRef.Id}' />
                                </filter>
                              </entity>
                            </fetch>";
                            EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (_entCol.Entities.Count > 0)
                                throw new InvalidPluginExecutionException("Duplicate Spare Part is not allowed.");
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
