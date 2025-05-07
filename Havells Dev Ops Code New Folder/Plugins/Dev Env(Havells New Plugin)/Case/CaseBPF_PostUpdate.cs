using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    public class CaseBPF_PostUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService.Trace("on the top");
            #endregion
            try
            {
                if (context.InputParameters.Contains("EntityMoniker") &&
                context.InputParameters["EntityMoniker"] is EntityReference && context.Depth <= 4)
                {
                    var entityRef = (EntityReference)context.InputParameters["EntityMoniker"];
                    var state = (OptionSetValue)context.InputParameters["State"];
                    var status = (OptionSetValue)context.InputParameters["Status"];
                    tracingService.Trace("Depth" + context.Depth);
                    tracingService.Trace("Status" + status.Value);
                    tracingService.Trace("State" + state.Value);
                    Entity entity = service.Retrieve("incident", entityRef.Id, new ColumnSet(true));
                    tracingService.Trace("On the top of entity");
                    tracingService.Trace(entity.LogicalName);
                    string logicalNameOfBPF = "phonetocaseprocess";
                    Entity activeProcessInstance = null;
                    int caseType = entity.Contains("casetypecode") ? entity.GetAttributeValue<OptionSetValue>("casetypecode").Value : 0;
                    EntityReference _entDepartment = entity.GetAttributeValue<EntityReference>("hil_casedepartment");
                    int statecode = entity.GetAttributeValue<OptionSetValue>("statecode").Value;
                    tracingService.Trace("case statecode value " + statecode);
                    if ((_entDepartment.Id == CaseConstaints._samparkDepartment) && (caseType == 2 || caseType == 3) && statecode == 1) // Check if case is being resolved
                    {
                        tracingService.Trace(statecode.ToString() + caseType.ToString());
                        activeProcessInstance = CasePostUpdate.GetActiveBPFDetails(entity, service);
                        if (activeProcessInstance != null)
                        {
                            tracingService.Trace("Entered in the if condition");
                            Guid activeBPFId = activeProcessInstance.Id;
                            Guid activeStageId = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());
                            int currentStagePosition = -1;
                            RetrieveActivePathResponse pathResp = CasePostUpdate.GetAllStagesOfSelectedBPF(activeBPFId, activeStageId, ref currentStagePosition, service);
                            if (currentStagePosition > -1 && pathResp.ProcessStages != null && pathResp.ProcessStages.Entities != null && currentStagePosition + 1 < pathResp.ProcessStages.Entities.Count)
                            {
                                tracingService.Trace("Entered in the 2nd if condition");
                                Guid nextStageId = (Guid)pathResp.ProcessStages.Entities[pathResp.ProcessStages.Entities.Count - 1].Attributes["processstageid"];
                                Entity entBPF = new Entity(logicalNameOfBPF)
                                {
                                    Id = activeBPFId
                                };
                                entBPF["activestageid"] = new EntityReference("processstage", nextStageId);
                                //entBPF["statuscode"] = new OptionSetValue(2);
                                service.Update(entBPF);
                                tracingService.Trace("last step");  
                                tracingService.Trace("Updated the entity");

                                var stateRequest = new SetStateRequest
                                {
                                    EntityMoniker = new EntityReference("phonetocaseprocess", activeProcessInstance.Id),
                                    State = new OptionSetValue(1), // Inactive.
                                    Status = new OptionSetValue(2) // Finished.

                                };
                                service.Execute(stateRequest);
                                tracingService.Trace("Marked as resolved");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Exception: {0}", ex.ToString());
                throw;
            }

        }
    }
}














