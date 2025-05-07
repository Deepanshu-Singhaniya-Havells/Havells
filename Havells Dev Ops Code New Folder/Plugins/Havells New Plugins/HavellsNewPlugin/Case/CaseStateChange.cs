using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel;

namespace HavellsNewPlugin.Case
{
    public class CaseStateChange : IPlugin
    {
        private static readonly Guid SamparkDepartmentId = new Guid("7bf1705a-3764-ee11-8df0-6045bdaa91c3");//Sampark Call center
        private static readonly Guid JobSubStatus = new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"); //Closed
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            //try
            //{
            if (context.InputParameters.Contains("EntityMoniker") &&
                context.InputParameters["EntityMoniker"] is EntityReference)
            {
                var entityRef = (EntityReference)context.InputParameters["EntityMoniker"];
                var state = (OptionSetValue)context.InputParameters["State"];
                var status = (OptionSetValue)context.InputParameters["Status"];

                Entity Case = service.Retrieve("incident", entityRef.Id, new ColumnSet("hil_casecategory", "hil_casedepartment", "hil_assignmentmatrix", "ownerid", "hil_job", "ticketnumber", "title", "adx_resolutiondate", "hil_assignmentlevel", "hil_pendinglevel"));
                Guid DepartmentId = Case.GetAttributeValue<EntityReference>("hil_casedepartment").Id;
                if (DepartmentId != SamparkDepartmentId)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_grievancehandlingactivity'>
                            <attribute name='activityid' />
                            <attribute name='subject' />
                            <attribute name='createdon' />
                            <attribute name='ownerid' />
                            <attribute name=""hil_caseassignmentmatrixlineid"" />
                            <order attribute='subject' descending='false' />
                            <filter type='and'>
                                <condition attribute='regardingobjectid' operator='eq' value='{entityRef.Id}' />
                                <condition attribute=""statecode"" operator=""in"">
                                <value>0</value>
                                <value>2</value>
                                </condition>
                            </filter>
                            </entity>
                        </fetch>";
                    EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entityCollection.Entities.Count > 0)
                    {
                        HashSet<string> _owners = new HashSet<string>();
                        string _activityOwner = string.Empty;
                        string _ownerName = string.Empty;
                        foreach (Entity ent in entityCollection.Entities)
                        {
                            _ownerName = ent.GetAttributeValue<EntityReference>("ownerid").Name;
                            if (!_owners.Contains(_ownerName))
                            {
                                _owners.Add(_ownerName);
                            }
                        }
                        foreach (string _owner in _owners)
                        {
                            _activityOwner += _owner + ", ";
                        }
                        if (_activityOwner != string.Empty)
                        {
                            _activityOwner = _activityOwner.Substring(0, _activityOwner.Length - 2);
                        }
                        if (entityCollection.Entities.Count == 1)
                        {
                            throw new InvalidPluginExecutionException(string.Format("There is {0} activitie open related to this grievance. Please ask {1} to complete pending activities first.", entityCollection.Entities.Count, _activityOwner));
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException(string.Format("There are {0} activitie(s) open related to this grievance. Please ask {1} to complete pending activities first.", entityCollection.Entities.Count, _activityOwner));
                        }
                    }

                    Guid JobId = Guid.Empty;

                    //if (Case.Contains("hil_job"))
                    //{
                    //    JobId = Case.GetAttributeValue<EntityReference>("hil_job").Id;

                    //    Entity Job = service.Retrieve("msdyn_workorder", JobId, new ColumnSet("msdyn_substatus", "msdyn_name"));
                    //    if (Job.Contains("msdyn_substatus"))
                    //    {
                    //        Guid susStatusId = Job.GetAttributeValue<EntityReference>("msdyn_substatus").Id;
                    //        string jobNumber = Job.Contains("msdyn_name") ? Job.GetAttributeValue<string>("msdyn_name") : "No job number found";
                    //        if (susStatusId != JobSubStatus)
                    //        {
                    //            throw new InvalidPluginExecutionException($"Corresponding Job({jobNumber}) should be closed in order to close the case");
                    //        }
                    //    }
                    //}

                    Guid AssignmentMatrixId = Case.GetAttributeValue<EntityReference>("hil_assignmentmatrix").Id;
                    Entity AssignmentMatrix = service.Retrieve("hil_caseassignmentmatrix", AssignmentMatrixId, new ColumnSet("hil_resolveby", "hil_spoc"));
                    Guid CaseOwnerID = Case.GetAttributeValue<EntityReference>("ownerid").Id;
                    int ResolveBy = AssignmentMatrix.GetAttributeValue<OptionSetValue>("hil_resolveby").Value;

                    if (ResolveBy == 1) // SPOC
                    {
                        Guid SPOCID = AssignmentMatrix.GetAttributeValue<EntityReference>("hil_spoc").Id;
                        if (context.InitiatingUserId != SPOCID)
                        {
                            string _ownerName = service.Retrieve("systemuser", SPOCID, new ColumnSet("fullname")).GetAttributeValue<string>("fullname");
                            throw new InvalidPluginExecutionException(string.Format("Access Denied! You are not Authorized to resolve the case. Please ask {0} to complete pending activities first.", _ownerName));
                        }
                    }
                    else if (ResolveBy == 2) // Assignee
                    {
                        if (context.InitiatingUserId != CaseOwnerID)
                        {
                            string _ownerName = service.Retrieve("systemuser", CaseOwnerID, new ColumnSet("fullname")).GetAttributeValue<string>("fullname");
                            throw new InvalidPluginExecutionException(string.Format("Access Denied! You are not Authorized to resolve the case. Please ask {0} to complete pending activities first.", _ownerName));
                        }
                    }

                    // Mark the assignment and pending level on case as case resolved.
                    UpdateAssignementandPendingLevel(Case, service);

                    //Send Email to all Stake Holder on Resolutio of the Case.
                    SendEmailOnCaseResolution _sendEmail = new SendEmailOnCaseResolution(service);
                    _sendEmail.SendEmail(Case);

                    if (IsBPFCreatedForRecord(Case, service))
                    {
                        string logicalNameOfBPF = "phonetocaseprocess";
                        Entity activeProcessInstance = GetActiveBPFDetails(Case, service);
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
                    }
                }
            }
            //}
            //catch (InvalidPluginExecutionException)
            //{
            //    throw;
            //}
            //catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            //{
            //    throw new InvalidPluginExecutionException("An FaultException occurred in the SetStatePlugin plug-in odf Enginner Shift.", ex);
            //}
            //catch (Exception ex)
            //{
            //    throw new InvalidPluginExecutionException("Error : " + ex);
            //}
        }
        private bool IsBPFCreatedForRecord(Entity Case, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("phonetocaseprocess");
            query.ColumnSet = new ColumnSet("incidentid");
            query.Criteria.AddCondition("incidentid", ConditionOperator.Equal, Case.Id);
            EntityCollection bpfColl = service.RetrieveMultiple(query);

            if (bpfColl.Entities.Count > 0) return true;
            return false;
        }
        private void UpdateAssignementandPendingLevel(Entity incident, IOrganizationService serivce)
        {
            incident["hil_assignmentlevel"] = new OptionSetValue(5); // Case Resolved
            incident["hil_pendinglevel"] = new OptionSetValue(9); // Case Resolved
            serivce.Update(incident); 
        }

        private Entity GetActiveBPFDetails(Entity entity, IOrganizationService service)
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
        private RetrieveActivePathResponse GetAllStagesOfSelectedBPF(Guid activeBPFId, Guid activeStageId, ref int currentStagePosition, IOrganizationService service)
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