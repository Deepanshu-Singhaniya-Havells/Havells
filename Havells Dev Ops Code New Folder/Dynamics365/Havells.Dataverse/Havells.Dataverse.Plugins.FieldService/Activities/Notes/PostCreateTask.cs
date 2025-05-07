using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Activities.Notes
{
    public class PostCreateTask : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "task" && (context.MessageName.ToUpper() == "CREATE" && context.Depth == 1))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                try
                {
                    if (entity.Contains("regardingobjectid"))
                    {
                        EntityReference _entRefRegarding = entity.GetAttributeValue<EntityReference>("regardingobjectid");
                        if (_entRefRegarding.LogicalName.ToLower() == "msdyn_workorder")
                            ProcessRequest(_entRefRegarding, _tracingService, service);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException($"ERROR! {ex.Message}");
                }
            }
        }
        public void ProcessRequest(EntityReference _jobEntityRef, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_workorderid' />
                <filter type='and'>
                    <condition attribute='msdyn_workorderid' operator='eq' value='{_jobEntityRef.Id}' />
                    <condition attribute='msdyn_substatus' operator='not-in'>
                    <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                    <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                    </condition>
                </filter>
                </entity>
                </fetch>";
                EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_entCol.Entities.Count > 0)
                {
                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_wacta'>
                    <attribute name='hil_wactaid' />
                    <filter type='and'>
                        <condition attribute='hil_jobid' operator='eq' value='{_jobEntityRef.Id}' />
                        <condition attribute='hil_wastatusreason' operator='eq' value='1' />
                    </filter>
                    </entity>
                    </fetch>";

                    _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));

                    if (_entCol.Entities.Count > 0)
                    {
                        Entity _entWACTA = new Entity("msdyn_workorder", _jobEntityRef.Id);
                        _entWACTA["hil_reopenbyconsumer"] = true;
                        service.Update(_entWACTA);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
