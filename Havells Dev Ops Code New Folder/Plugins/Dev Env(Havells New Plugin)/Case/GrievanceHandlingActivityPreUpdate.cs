using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.AccessControl;

namespace HavellsNewPlugin.Case
{
    public class GrievanceHandlingActivityPreUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_grievancehandlingactivity" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    //tracingService.Trace(entity.LogicalName);
                    if (entity.Contains("statecode"))
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("statecode").Value == 2)//Cancelled
                        {
                            throw new InvalidPluginExecutionException("Access Denied !!! You are not authorized to Cancel the Activity.");
                        }
                        else if (entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)//Completed
                        {
                            #region Validations
                            Entity _entActivity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("ownerid", "regardingobjectid", "hil_activitytype", "hil_caseassignmentmatrixlineid"));
                            EntityReference owner = _entActivity.GetAttributeValue<EntityReference>("ownerid");

                            if (owner.Id != context.InitiatingUserId)
                            {
                                throw new InvalidPluginExecutionException("Access Denied !!! You are not authorized to update the Activity.");
                            }
                            //Checking for Open Activities
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_grievancehandlingactivity'>
                                <attribute name='subject' />
                                <order attribute='createdon' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_activitytype' operator='eq' value='1' />
                                  <condition attribute='statecode' operator='ne' value='1' />
                                  <condition attribute='regardingobjectid' operator='eq' value='{_entActivity.GetAttributeValue<EntityReference>("regardingobjectid").Id}' />
                                  <condition attribute='activityid' operator='ne' value='{entity.Id}' />
                                </filter>
                              </entity>
                            </fetch>";

                            EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (_entCol.Entities.Count > 0)
                            {
                                throw new InvalidPluginExecutionException("Access Denied !!! There are " + _entCol.Entities.Count.ToString() + " Open Activities found agasint the Case.");
                            }
                            #endregion
                            OptionSetValue _activityType = _entActivity.GetAttributeValue<OptionSetValue>("hil_activitytype");
                            if (_activityType.Value == CaseConstaints._activityType_Assignment)
                            {
                                EntityReference _entRefCase = _entActivity.GetAttributeValue<EntityReference>("regardingobjectid");
                                Entity _entCaseRecord = service.Retrieve(_entRefCase.LogicalName, _entRefCase.Id, new ColumnSet("hil_assignmentmatrix", "firstresponsesent"));

                                bool _firstResponseSent = _entCaseRecord.Contains("firstresponsesent") ? _entCaseRecord.GetAttributeValue<bool>("firstresponsesent") : false;

                                if (!_firstResponseSent) {
                                    throw new InvalidPluginExecutionException("Access Denied !!! We can't find First Response in the System. Please make a Phone call/Email first.");
                                }
                                Entity _entCase = new Entity(_entRefCase.LogicalName, _entRefCase.Id);
                                _entCase["hil_workdoneon"] = DateTime.Now.AddMinutes(330);
                                _entCase["hil_isworkdone"] = true;
                                service.Update(_entCase);

                                Entity assignmentmatrixRecord = service.Retrieve(_entCaseRecord.GetAttributeValue<EntityReference>("hil_assignmentmatrix").LogicalName, _entCaseRecord.GetAttributeValue<EntityReference>("hil_assignmentmatrix").Id, new ColumnSet("hil_resolveby", "hil_spoc", "hil_spocsla"));

                                int resolveBy = assignmentmatrixRecord.GetAttributeValue<OptionSetValue>("hil_resolveby").Value;
                                if (resolveBy == CaseConstaints._activityResolveBy_SPOC)
                                {
                                    int spocSla = assignmentmatrixRecord.GetAttributeValue<int>("hil_spocsla");
                                    Guid spocId = assignmentmatrixRecord.GetAttributeValue<EntityReference>("hil_spoc").Id;
                                    DateTime _caseResolveBy = DateTime.Now.AddMinutes(330 + spocSla);
                                    Entity _entAssignmentMatrixLine = new Entity(_entActivity.GetAttributeValue<EntityReference>("hil_caseassignmentmatrixlineid").LogicalName, _entActivity.GetAttributeValue<EntityReference>("hil_caseassignmentmatrixlineid").Id);
                                    CaseAssignmentEngine _caseAssignment = new CaseAssignmentEngine();
                                    _caseAssignment.CreateGrievanceHandlingActivity(service, _entCase, string.Empty, "Assignment to SPOC Case# ", _entAssignmentMatrixLine, _caseResolveBy, assignmentmatrixRecord.GetAttributeValue<EntityReference>("hil_spoc"), 4);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}
