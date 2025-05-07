using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace D365WebJobs
{
    public class GrivanceModule
    {
        public void CreateCallCDRREcord(IOrganizationService service, CDR_Request1 request) {
            string cdr_report = Newtonsoft.Json.JsonConvert.SerializeObject(request);
            Entity _entCallResponse = new Entity("hil_phonecallresponses");
            _entCallResponse["hil_name"] = request.Correlation_ID;
            _entCallResponse["hil_responsejson"] = cdr_report;
            service.Create(_entCallResponse);
        }
        public static void GrivanceModule_PostCreate(IOrganizationService service) {
            string _returnMsg = string.Empty;
            try
            {
                Entity entity = service.Retrieve("incident", new Guid("f916856a-4e75-44d4-a7c5-d562b50002e5"), new ColumnSet(true));

                Entity caseEntity = new Entity(entity.LogicalName, entity.Id);

                string caseNumber = entity.GetAttributeValue<string>("title");
                int caseType = entity.GetAttributeValue<OptionSetValue>("casetypecode").Value;
                EntityReference entDepartment = entity.GetAttributeValue<EntityReference>("hil_casedepartment");
                EntityReference entBranch = entity.GetAttributeValue<EntityReference>("hil_branch");
                string _fetchXML = string.Empty;
                EntityCollection _entColl = null;
                int _resolutionSLA = 0;
                int _firstResponseSLA = 0;

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
                        CreateGrievanceHandlingActivity(service, entity, caseNumber, "Assignment of Case# ", _entColl[0], _caseResolveBy);
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
                Console.WriteLine("Error : " + ex.Message);
            }
            Console.WriteLine(_returnMsg);
        }
        public static void CreateGrievanceHandlingActivity(IOrganizationService _service, Entity _caseEntity, string _caseNumber, string _subject, Entity _caseAssignmentMatrixLine, DateTime _caseResolveBy)
        {
            try
            {
                Entity _grievanceActivity = new Entity("hil_grievancehandlingactivity");

                if (string.IsNullOrWhiteSpace(_caseNumber))
                {
                    _caseNumber = _service.Retrieve(_caseEntity.LogicalName, _caseEntity.Id, new ColumnSet("title")).GetAttributeValue<string>("title");
                }
                _grievanceActivity["subject"] = _subject + _caseNumber;

                _grievanceActivity["regardingobjectid"] = _caseEntity.ToEntityReference();
                _grievanceActivity["hil_activitytype"] = new OptionSetValue(1);//Assignment
                _grievanceActivity["scheduledstart"] = DateTime.Now.AddMinutes(330);
                _grievanceActivity["scheduledend"] = _caseResolveBy;
                _grievanceActivity["hil_caseassignmentmatrixlineid"] = _caseAssignmentMatrixLine.ToEntityReference();
                _service.Create(_grievanceActivity);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        private static void AssignCase(IOrganizationService _service, Entity _entity, EntityReference _assignee, EntityReference _assignmentMatrix, DateTime _firstResponseBy, DateTime _caseResolveBy)
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
    public class CDR_Request1
    {
        public string Caller_Id { get; set; }
        public string Caller_Name { get; set; }
        public string Caller_Number { get; set; }
        public string Call_Type { get; set; }
        public string Caller_Status { get; set; }
        public string Conversation_Duration { get; set; }
        public string Correlation_ID { get; set; }
        public string Date { get; set; }
        public string Destination_Name { get; set; }
        public string Destination_Number { get; set; }
        public string Destination_Status { get; set; }
        public string Overall_Call_Duration { get; set; }
        public string Overall_Call_Status { get; set; }
        public string Recording { get; set; }
        public string Time { get; set; }
    }
}
