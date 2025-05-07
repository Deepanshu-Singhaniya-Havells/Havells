using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GrievanceHandling
{
    public class Escalations
    {
        private readonly IOrganizationService service;
        internal Escalations(IOrganizationService _service)
        {
            service = _service;
        }
        private void UpdateActivity(Entity grievanceActivity)
        {
            grievanceActivity["hil_isescalated"] = true;
            service.Update(grievanceActivity);
        }
        private void UpdateCase(Entity Case, DateTime resolveBy, int level)
        {
            Case["isescalated"] = true;
            Case["escalatedon"] = DateTime.Now.AddMinutes(330);


            //Update the assignment level on the case. 
            if (level == 2 || level == 3)
            {
                Case["hil_assignmentlevel"] = new OptionSetValue(level);
            }
            else
            {
                Case["hil_assignmentlevel"] = new OptionSetValue(4);// Escalated
            }


            TimeSpan _diffMin = resolveBy - DateTime.Now.AddMinutes(330);
            if (_diffMin.TotalMinutes <= 60)
            {
                Case["hil_resolvebyduration"] = Convert.ToInt32(_diffMin.TotalMinutes) + " Minutes";
            }
            else if (_diffMin.TotalMinutes < 1440)
            {
                double _hrs = _diffMin.TotalMinutes / 60.0;
                if (Convert.ToInt32(_diffMin.TotalMinutes) % 60 != 0)
                {
                    Case["hil_resolvebyduration"] = Convert.ToInt32(_diffMin.TotalMinutes) / 60 + 1 + " hrs";
                }
                else
                    Case["hil_resolvebyduration"] = Math.Round(_hrs, 0) + " hrs";
            }
            else
            {
                int _hr = Convert.ToInt32(_diffMin.TotalMinutes) / 1440;
                Case["hil_resolvebyduration"] = _hr + " days";
            }

            service.Update(Case);
        }
        internal EntityCollection FetchCaseAssignmentMatrixLine(Guid matrixLineId)
        {
            Entity matrixLine = service.Retrieve("hil_caseassignmentmatrixline", matrixLineId, new ColumnSet("hil_caseassignmentmatrixid", "hil_level"));
            int currentLevel = matrixLine.Contains("hil_level") ? matrixLine.GetAttributeValue<int>("hil_level") : 0;
            Guid matrixID = matrixLine.GetAttributeValue<EntityReference>("hil_caseassignmentmatrixid").Id;
            QueryExpression query = new QueryExpression("hil_caseassignmentmatrixline");
            query.ColumnSet = new ColumnSet("hil_sla", "hil_level", "hil_name", "hil_assigneeuser", "hil_assigneeteam");
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("hil_caseassignmentmatrixid", ConditionOperator.Equal, matrixID);
            query.Criteria.AddCondition("hil_level", ConditionOperator.Equal, currentLevel + 1);
            return service.RetrieveMultiple(query);

        }
        private void CreateGrievanceHandlingActivity(IOrganizationService _service, Entity _caseEntity, string _subject, Entity _caseAssignmentMatrixLine, DateTime _caseResolveBy, EntityReference _assignee, int _activityType)
        {
            try
            {
                Entity _grievanceActivity = new Entity("hil_grievancehandlingactivity");
                _grievanceActivity["subject"] = _subject;
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
        internal void EscalateIncident()
        {
            EntityCollection expiredColl = OverDueActivities();
            if (expiredColl.Entities.Count == 0) { Console.WriteLine("No Overdue Grievance Activity found."); return; }
            Console.WriteLine($"Total Case Count to be processed... {expiredColl.Entities.Count}");
            foreach (Entity entity in expiredColl.Entities)
            {
                try
                {
                    string subject = entity.Contains("subject") ? entity.GetAttributeValue<string>("subject") : " ";
                    Console.WriteLine($"Processing... Case# {entity.GetAttributeValue<EntityReference>("regardingobjectid")}");
                    Guid CaseId = entity.GetAttributeValue<EntityReference>("regardingobjectid").Id;

                    Guid matrixLineId = entity.GetAttributeValue<EntityReference>("hil_caseassignmentmatrixlineid").Id;
                    Entity Case = service.Retrieve("incident", CaseId, new ColumnSet("isescalated", "escalatedon", "hil_resolvebyduration", "hil_assignmentlevel"));

                    EntityCollection matrixLinesColl = FetchCaseAssignmentMatrixLine(matrixLineId);
                    if (matrixLinesColl.Entities.Count > 0)
                    {
                        int level = matrixLinesColl.Entities[0].Contains("hil_level") ? matrixLinesColl.Entities[0].GetAttributeValue<int>("hil_level") : -1;
                        int SLA = matrixLinesColl.Entities[0].Contains("hil_sla") ? matrixLinesColl.Entities[0].GetAttributeValue<int>("hil_sla") : 0;
                        DateTime resolveBy = DateTime.Now.AddMinutes(330 + SLA);
                        EntityReference _assignee = matrixLinesColl.Entities[0].Contains("hil_assigneeuser") ? matrixLinesColl.Entities[0].GetAttributeValue<EntityReference>("hil_assigneeuser") : matrixLinesColl.Entities[0].GetAttributeValue<EntityReference>("hil_assigneeteam");
                        CreateGrievanceHandlingActivity(service, Case, subject, matrixLinesColl.Entities[0], resolveBy, _assignee, 2);
                        UpdateActivity(entity);
                        UpdateCase(Case, resolveBy, level);
                    }
                    else
                    {
                        Console.WriteLine("There are no further Esclation Matrix lines");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Processing... Case# {entity.GetAttributeValue<EntityReference>("regardingobjectid")} ERROR!!! " + ex.Message);
                }
            }
        }
        private EntityCollection OverDueActivities()
        {
            //find the activities that has expired in the past 1 hour, and escalate them to new level; 
            QueryExpression query = new QueryExpression("hil_grievancehandlingactivity");
            query.ColumnSet = new ColumnSet("regardingobjectid", "hil_caseassignmentmatrixlineid", "scheduledstart", "scheduledend", "ownerid", "hil_activitytype", "subject", "statecode", "hil_isescalated");
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("hil_activitytype", ConditionOperator.In, new object[] { 1, 2 });//assignment,Escalations
            query.Criteria.AddCondition("scheduledend", ConditionOperator.OlderThanXMinutes, 1);
            query.Criteria.AddCondition("hil_isescalated", ConditionOperator.Equal, false);
            query.Criteria.AddCondition("hil_caseassignmentmatrixlineid", ConditionOperator.NotNull);
            return service.RetrieveMultiple(query);
        }
    }
}