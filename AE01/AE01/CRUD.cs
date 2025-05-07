using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

internal class CRUD
{
    private readonly IOrganizationService service;

    public CRUD(IOrganizationService _service)
    {
        service = _service;
    }
    public async Task CreateRecording(CDR_Request req, Guid ActivityId)
    {
        EntityCollection tempCollection = await checkExistingRecording(ActivityId);
        if (tempCollection.Entities.Count == 0)
        {
            Entity recording = new Entity("msdyn_recording");
            recording["msdyn_ci_url"] = req.Recording;
            recording["msdyn_ci_transcript_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(req);
            recording["msdyn_phone_call_activity"] = new EntityReference("phonecall", ActivityId);
            service.Create(recording);
        }
        else
        {
            tempCollection.Entities[0]["msdyn_ci_url"] = req.Recording;
            tempCollection.Entities[0]["msdyn_ci_transcript_json"] = Newtonsoft.Json.JsonConvert.SerializeObject(req);
            tempCollection.Entities[0]["msdyn_phone_call_activity"] = new EntityReference("phonecall", ActivityId);
            service.Update(tempCollection.Entities[0]);
        }
    }

    public async Task<EntityCollection> checkExistingRecording(Guid ActivityID)
    {
        QueryExpression query = new QueryExpression("msdyn_recording");
        query.ColumnSet = new ColumnSet("msdyn_ci_url", "msdyn_ci_transcript_json", "msdyn_phone_call_activity");
        query.Criteria.AddCondition("msdyn_recording", ConditionOperator.Equal, ActivityID);
        return service.RetrieveMultiple(query);
    }

    public async Task<CDR_Response> pushCDR(CDR_Request req)
    {
        CDR_Response response = new CDR_Response();

        response.status = "Success";
        return response;
    }


    public class CDR_Response
    {
        public string? status { get; set; }

    }


    public class CDR_Request
    {
        public string? Caller_Id { get; set; }
        public string? Caller_Name { get; set; }
        public string? Caller_Number { get; set; }
        public string? Call_Type { get; set; }
        public string? Caller_Status { get; set; }
        public string? Conversation_Duration { get; set; }
        public string? Correlation_ID { get; set; }
        public string? Date { get; set; }
        public string? Destination_Name { get; set; }
        public string? Destination_Number { get; set; }
        public string? Destination_Status { get; set; }
        public string? Overall_Call_Duration { get; set; }
        public string? Overall_Call_Status { get; set; }
        public string? Recording { get; set; }
        public string? Time { get; set; }

    }

}
