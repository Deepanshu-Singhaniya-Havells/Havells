using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Service_Call
{
    public class CancelServiceJob : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string JsonResponse = "";
            Guid JobGuid = Guid.Empty;
            bool isValidJobGuid = false;

            string JobId = Convert.ToString(context.InputParameters["JobGuid"]);
            string Source = Convert.ToString(context.InputParameters["Source"]);
            string JobNumber = Convert.ToString(context.InputParameters["JobNumber"]);

            if (string.IsNullOrWhiteSpace(Source))
            {
                if (string.IsNullOrWhiteSpace(JobId))
                {
                    JsonResponse = JsonSerializer.Serialize(new CancelJobResponse
                    {
                        StatusCode = "204",
                        StatusDescription = "Job Guid is required."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                isValidJobGuid = Guid.TryParse(JobId, out JobGuid);
                if (!isValidJobGuid || JobGuid == Guid.Empty)
                {
                    JsonResponse = JsonSerializer.Serialize(new CancelJobResponse
                    {
                        StatusCode = "204",
                        StatusDescription = "Invalid Job Guid."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                JsonResponse = JsonSerializer.Serialize(ToCancelServiceJob(service, JobGuid.ToString(), ""));
                _tracingService.Trace(JsonResponse);
                context.OutputParameters["data"] = JsonResponse;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(JobNumber))
                {
                    JsonResponse = JsonSerializer.Serialize(new CancelJobResponse
                    {
                        StatusCode = "204",
                        StatusDescription = "Job Number is required."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (Source == "12")
                {
                    JsonResponse = JsonSerializer.Serialize(ToCancelServiceJob(service, JobNumber, Source));
                    _tracingService.Trace(JsonResponse);
                    context.OutputParameters["data"] = JsonResponse;
                }
                else
                {
                    JsonResponse = JsonSerializer.Serialize(new CancelJobResponse
                    {
                        StatusCode = "204",
                        StatusDescription = "Invalid Source."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
        }
        public CancelJobResponse ToCancelServiceJob(IOrganizationService service, string JobId, string Source)
        {
            CancelJobResponse res = new CancelJobResponse();
            Guid JobGuid = Guid.Empty;
            try
            {
                if (service != null)
                {
                    string _cancelStatusId = "1527FA6C-FA0F-E911-A94E-000D3AF060A1";
                    if (!string.IsNullOrWhiteSpace(Source))
                    {
                        QueryExpression Qry = new QueryExpression("msdyn_workorder");
                        Qry.ColumnSet = new ColumnSet(false);
                        Qry.Criteria = new FilterExpression(LogicalOperator.And);
                        Qry.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, JobId);
                        EntityCollection entColl = service.RetrieveMultiple(Qry);
                        if (entColl.Entities.Count > 0)
                        {
                            JobGuid = entColl.Entities[0].Id;
                        }
                        else
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Invalid Job Number." };
                            return res;
                        }
                    }
                    else
                    {
                        JobGuid = new Guid(JobId);
                    }
                    Entity entWorkOrder = service.Retrieve("msdyn_workorder", JobGuid, new ColumnSet("msdyn_substatus"));
                    if (entWorkOrder != null)
                    {
                        EntityReference _erSubstatus = entWorkOrder.GetAttributeValue<EntityReference>("msdyn_substatus");
                        if (_erSubstatus.Id.ToString().ToUpper() == _cancelStatusId)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job is already Cancelled" };
                            return res;
                        }
                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_statustransitionmatrix'>
                        <attribute name='hil_statustransitionmatrixid' />
                        <filter type='and'>
                            <condition attribute='hil_basestatus' operator='eq' value='" + _erSubstatus.Id + @"' />
                            <filter type='or'>
                            <condition attribute='hil_statustransition1' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition2' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition3' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition4' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition5' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition6' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition7' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition8' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition9' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition10' operator='eq' value='{" + _cancelStatusId + @"}' />
                            </filter>
                        </filter>
                        </entity>
                        </fetch>";
                        EntityCollection entCols = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCols.Entities.Count == 0)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job cannot be Cancelled, please contact with Customer Care." };
                            return res;
                        }
                    }
                    else
                    {
                        res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Invalid Job ID." };
                        return res;
                    }

                    QueryExpression Query = new QueryExpression("msdyn_workordersubstatus");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, "Canceled");
                    EntityCollection entColls = service.RetrieveMultiple(Query);
                    if (entColls.Entities.Count > 0)
                    {
                        Entity _jobUpdate = new Entity("msdyn_workorder", JobGuid);
                        switch (Source)
                        {
                            case "12":
                                _jobUpdate["hil_closureremarks"] = "Job cancelled by Customer via Whats App";
                                break;
                            default:
                                _jobUpdate["hil_closureremarks"] = "Job cancelled by Customer via One App";
                                break;
                        }
                        _jobUpdate["msdyn_substatus"] = new EntityReference(entColls.EntityName, entColls[0].Id);

                        DateTime _today = DateTime.Now.AddMinutes(330);
                        _jobUpdate["hil_jobclosureon"] = _today;
                        _jobUpdate["msdyn_timeclosed"] = _today;

                        try
                        {
                            service.Update(_jobUpdate);
                        }
                        catch (Exception ex)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job cannot be Cancelled, please contact with Customer Care. \n" + ex.Message };
                            return res;
                        }
                        res.StatusCode = "200";
                        res.StatusDescription = "Job is Cancelled";
                    }
                    else
                    {
                        res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Job Sub status(Cancel) is not defined in System." };
                        return res;
                    }
                }
                else
                {
                    res = new CancelJobResponse { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                res = new CancelJobResponse { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return res;
        }
    }
    public class CancelJobResponse
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class CancelJobRequest
    {
        public string JobGuid { get; set; }
    }
}
