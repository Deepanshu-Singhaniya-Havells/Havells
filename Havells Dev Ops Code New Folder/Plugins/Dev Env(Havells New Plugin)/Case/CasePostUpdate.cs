using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.Security.AccessControl;

namespace HavellsNewPlugin.Case
{
    public class CasePostUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider) 
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("on the top");
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                  && context.PrimaryEntityName.ToLower() == "incident" && context.Depth == 1)
                {             
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    tracingService.Trace(entity.LogicalName);
                    string logicalNameOfBPF = "phonetocaseprocess";
                    Entity activeProcessInstance = null;              
                    int caseType = entity.Contains("casetypecode") ? entity.GetAttributeValue<OptionSetValue>("casetypecode").Value : 0;
                    EntityReference _entDepartment = entity.GetAttributeValue<EntityReference>("hil_casedepartment");
                    int caseStatus = entity.GetAttributeValue<OptionSetValue>("statecode").Value;
                    if (entity.Contains("hil_isocr") && entity.GetAttributeValue<bool>("hil_isocr") && _entDepartment.Id == CaseConstaints._samparkDepartment)
                    {
                        if (caseType == 1)
                        {
                            activeProcessInstance = GetActiveBPFDetails(entity, service);
                            if (activeProcessInstance != null)
                            {
                                Guid activeBPFId = activeProcessInstance.Id; // Id of the active process instance, which will be used
                                                                             // Retrieve the active stage ID of in the active process instance
                                Guid activeStageId = new Guid(activeProcessInstance.Attributes["processstageid"].ToString());
                                int currentStagePosition = -1;
                                RetrieveActivePathResponse pathResp = GetAllStagesOfSelectedBPF(activeBPFId, activeStageId, ref currentStagePosition, service);
                                if (currentStagePosition > -1 && pathResp.ProcessStages != null && pathResp.ProcessStages.Entities != null && currentStagePosition + 1 < pathResp.ProcessStages.Entities.Count)
                                {
                                    // Retrieve the stage ID of the next stage that you want to set as active
                                    Guid nextStageId = (Guid)pathResp.ProcessStages.Entities[pathResp.ProcessStages.Entities.Count - 1].Attributes["processstageid"];
                                    // Set the next stage as the active stage
                                    Entity entBPF = new Entity(logicalNameOfBPF)
                                    {
                                        Id = activeBPFId
                                    };
                                    entBPF["activestageid"] = new EntityReference("processstage", nextStageId);
                                    service.Update(entBPF);
                                    var stateRequest = new SetStateRequest
                                    {
                                        EntityMoniker = new EntityReference("phonetocaseprocess", activeProcessInstance.Id),
                                        State = new OptionSetValue(1), // Inactive.
                                        Status = new OptionSetValue(2) // Finished.
                                    };
                                    service.Execute(stateRequest);
                                }
                            }
                            var incidentResolution = new Entity("incidentresolution");
                            incidentResolution["subject"] = "Information Provided";
                            incidentResolution["incidentid"] = entity.ToEntityReference();
                            incidentResolution["description"] = entity.GetAttributeValue<string>("adx_resolution");
                            //incidentResolution["timespent"] = 60;

                            var closeIncidentRequest = new CloseIncidentRequest
                            {
                                IncidentResolution = incidentResolution,
                                Status = new OptionSetValue(1000)
                            };
                            service.Execute(closeIncidentRequest);
                        }
                    }                  
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
        public static Entity GetActiveBPFDetails(Entity entity, IOrganizationService service)
        {
            Entity activeProcessInstance = null;
            RetrieveProcessInstancesRequest entityBPFsRequest = new RetrieveProcessInstancesRequest
            {
                EntityId = entity.Id,
                EntityLogicalName = entity.LogicalName
            };
            RetrieveProcessInstancesResponse entityBPFsResponse = (RetrieveProcessInstancesResponse)service.Execute(entityBPFsRequest);
            if (entityBPFsResponse.Processes != null && entityBPFsResponse.Processes.Entities != null)
            {
                activeProcessInstance = entityBPFsResponse.Processes.Entities[0];
            }
            return activeProcessInstance;
        }
        public static RetrieveActivePathResponse GetAllStagesOfSelectedBPF(Guid activeBPFId, Guid activeStageId, ref int currentStagePosition, IOrganizationService service)
        {
            // Retrieve the process stages in the active path of the current process instance
            RetrieveActivePathRequest pathReq = new RetrieveActivePathRequest
            {
                ProcessInstanceId = activeBPFId
            };
            RetrieveActivePathResponse pathResp = (RetrieveActivePathResponse)service.Execute(pathReq);
            for (int i = 0; i < pathResp.ProcessStages.Entities.Count; i++)
            {
                // Retrieve the active stage name and active stage position based on the activeStageId for the process instance
                if (pathResp.ProcessStages.Entities[i].Attributes["processstageid"].ToString() == activeStageId.ToString())
                {
                    currentStagePosition = i;
                }
            }
            return pathResp;
        }
    }
}