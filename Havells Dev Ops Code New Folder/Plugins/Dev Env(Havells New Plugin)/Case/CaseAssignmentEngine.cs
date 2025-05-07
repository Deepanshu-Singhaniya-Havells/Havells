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
    class CaseAssignmentEngine
    {
        public string AssignCase(IOrganizationService service, Entity _caseEntity)
        {
            string _returnMsg = "DONE";
            try
            {
                Entity entity = service.Retrieve(_caseEntity.LogicalName, _caseEntity.Id, new ColumnSet(true));
                Entity caseEntity = new Entity(entity.LogicalName, entity.Id);
                string caseNumber = entity.GetAttributeValue<string>("title");
                int caseType = entity.GetAttributeValue<OptionSetValue>("casetypecode").Value;
                EntityReference entDepartment = entity.GetAttributeValue<EntityReference>("hil_casedepartment");
                EntityReference entBranch = entity.GetAttributeValue<EntityReference>("hil_branch");
                EntityReference CaseCategory = entity.GetAttributeValue<EntityReference>("hil_casecategory");
                int Origin = entity.Contains("caseorigincode") ? entity.GetAttributeValue<OptionSetValue>("caseorigincode").Value : 0;
                string _fetchXML = string.Empty;
                EntityCollection _entColl = null;
                int _resolutionSLA = 0;
                int _firstResponseSLA = 0;

                if (entDepartment.Id == CaseConstaints._samparkDepartment)//Sampark Call Center and Case Type !=Enquiry. For Ennquiry we are resolving Case On Call.
                {
                    if (caseType != 1)
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_caseassignmentmatrixline'>
                                <attribute name='hil_assigneeuser' />
                                <attribute name='hil_assigneeteam' />
                                <attribute name='hil_caseassignmentmatrixid' />
                                <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_level' operator='eq' value='1' />
                                    <filter type='or'>
                                        <condition attribute='hil_assigneeuser' operator='not-null' />
                                        <condition attribute='hil_assigneeteam' operator='not-null' />
                                    </filter>
                                </filter>
                                <link-entity name='hil_caseassignmentmatrix' from='hil_caseassignmentmatrixid' to='hil_caseassignmentmatrixid' link-type='inner' alias='ac'>
                                  <filter type='and'>
                                    <condition attribute='hil_department' operator='eq' value='{entDepartment.Id}' />
                                    <condition attribute='hil_branch' operator='eq' value='{entBranch.Id}' />
                                     <condition attribute='hil_caseorigin' operator='eq' value='{Origin}' />
                                    <condition attribute='hil_casetype' operator='eq' value='{caseType}' />
                                     <condition attribute='hil_casecategory' operator='eq' value='{CaseCategory.Id}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
                        _entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entColl.Entities.Count > 0)
                        {
                            DateTime _firstResponseBy = DateTime.Now.AddHours(_firstResponseSLA);
                            DateTime _caseResolveBy = DateTime.Now.AddHours(_resolutionSLA);

                            EntityReference _assignmentMatrix = _entColl[0].GetAttributeValue<EntityReference>("hil_caseassignmentmatrixid");
                            EntityReference _assignee = _entColl[0].Contains("hil_assigneeuser") ? _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeuser") : _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                            AssignCase(service, caseEntity, _assignee, _assignmentMatrix, _firstResponseBy, _caseResolveBy);
                            _returnMsg = "Case has been assigned to concern person.";
                        }
                        else
                        {
                            _returnMsg = "Case Assignment Matrix doesn't exist in System.";
                        }
                    }
                    else
                    {
                        _returnMsg = "DONE";
                    }
                }
                else
                {
                    int caseOrigin = entity.GetAttributeValue<OptionSetValue>("caseorigincode").Value;
                    EntityReference entCaseCategory = entity.GetAttributeValue<EntityReference>("hil_casecategory");
                    EntityReference entCaseProduct = entity.GetAttributeValue<EntityReference>("productid");

                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_caseassignmentmatrixline'>
                        <attribute name='hil_caseassignmentmatrixlineid' />
                        <attribute name='hil_caseassignmentmatrixid' />
                        <attribute name='hil_assigneeuser' />
                        <attribute name='hil_assigneeteam' />
                        <attribute name='hil_sla' />
                        <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_level' operator='eq' value='1' />
                                <filter type='or'>
                                    <condition attribute='hil_assigneeuser' operator='not-null' />
                                    <condition attribute='hil_assigneeteam' operator='not-null' />
                                </filter>
                            </filter>
                        <link-entity name='hil_caseassignmentmatrix' from='hil_caseassignmentmatrixid' to='hil_caseassignmentmatrixid' link-type='inner' alias='af'>
                        <attribute name='hil_firstresponsesla' />
                        <filter type='and'>
                        <condition attribute='hil_department' operator='eq' value='{entDepartment.Id}' />
                        <condition attribute='hil_branch' operator='eq' value='{entBranch.Id}' />
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_caseorigin' operator='eq' value='{caseOrigin}' />
                        <condition attribute='hil_casetype' operator='eq' value='{caseType}' />
                        <condition attribute='hil_casecategory' operator='eq' value='{entCaseCategory.Id}' />";
                    if (entCaseProduct != null)
                        _fetchXML = _fetchXML + $@"<condition attribute='hil_productdivision' operator='eq' value='{entCaseProduct.Id}' />";
                    _fetchXML = _fetchXML + $@"</filter>
                        </link-entity>
                        </entity>
                        </fetch>";
                    _entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (_entColl.Entities.Count > 0)
                    {
                        if (_entColl[0].Contains("hil_sla"))
                            _resolutionSLA = _entColl[0].GetAttributeValue<int>("hil_sla");
                        if (_entColl[0].Contains("af.hil_firstresponsesla"))
                            _firstResponseSLA = Convert.ToInt32(_entColl[0].GetAttributeValue<AliasedValue>("af.hil_firstresponsesla").Value.ToString());
                        int _baseMinutes = 330;
                        DateTime _firstResponseBy = DateTime.Now.AddMinutes(_baseMinutes + _firstResponseSLA);
                        DateTime _caseResolveBy = DateTime.Now.AddMinutes(_baseMinutes + _resolutionSLA);

                        EntityReference _assignmentMatrix = _entColl[0].GetAttributeValue<EntityReference>("hil_caseassignmentmatrixid");
                        EntityReference _assignee = _entColl[0].Contains("hil_assigneeuser") ? _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeuser") : _entColl[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                        AssignCase(service, caseEntity, _assignee, _assignmentMatrix, _firstResponseBy, _caseResolveBy);
                        CreateGrievanceHandlingActivity(service, entity, caseNumber, "Assignment of Case# ", _entColl[0], _caseResolveBy, _assignee, 1);
                        _returnMsg = "Case has been assigned to concern person.";
                    }
                    else
                    {
                        _returnMsg = "Case Assignment Matrix doesn't exist in System.";
                    }
                }
            }
            catch (Exception ex)
            {
                _returnMsg = "Error : " + ex.Message;
            }
            return _returnMsg;
        }
        public void CreateGrievanceHandlingActivity(IOrganizationService _service, Entity _caseEntity, string _caseNumber, string _subject, Entity _caseAssignmentMatrixLine, DateTime _caseResolveBy, EntityReference _assignee, int _activityType)
        {
            try
            {
                Entity _grievanceActivity = new Entity("hil_grievancehandlingactivity");

                if (string.IsNullOrWhiteSpace(_caseNumber))
                {
                    _caseNumber = _service.Retrieve(_caseEntity.LogicalName, _caseEntity.Id, new ColumnSet("title")).GetAttributeValue<string>("title");
                }
                _grievanceActivity["subject"] = "Case ID " + _caseNumber + " regarding concern 'Delay in Service' has been assigned to you. Please resolve by " + _caseResolveBy.ToString("dd-MM-yyyy HH:mm:ss");//_subject + _caseNumber;

                _grievanceActivity["regardingobjectid"] = _caseEntity.ToEntityReference();
                _grievanceActivity["hil_activitytype"] = new OptionSetValue(_activityType);//Assignment
                _grievanceActivity["scheduledstart"] = DateTime.Now.AddMinutes(330);
                _grievanceActivity["scheduledend"] = _caseResolveBy;
                _grievanceActivity["ownerid"] = _assignee;
                _grievanceActivity["hil_caseassignmentmatrixlineid"] = _caseAssignmentMatrixLine.ToEntityReference();
                _service.Create(_grievanceActivity);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private void AssignCase(IOrganizationService _service, Entity _entity, EntityReference _assignee, EntityReference _assignmentMatrix, DateTime _firstResponseBy, DateTime _caseResolveBy)
        {
            try
            {
                Entity _entCase = new Entity(_entity.LogicalName, _entity.Id);
                _entCase["hil_assignmentmatrix"] = _assignmentMatrix;
                _entCase["responseby"] = _firstResponseBy;
                _entCase["followupby"] = _firstResponseBy;
                _entCase["resolveby"] = _caseResolveBy;
                _entCase["ownerid"] = _assignee;
                _entCase["hil_caseassignedon"] = DateTime.Now.AddMinutes(330);
                _service.Update(_entCase);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }

}
