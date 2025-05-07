using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

internal class FetchJobsSheetData(IOrganizationService _service)
{
    private QueryExpression query = new();
    private readonly IOrganizationService service = _service;

    internal List<Entity> FetchJobIncidentData(Guid JobId)
    {
        query = new QueryExpression("msdyn_workorderincident");
        query.ColumnSet = new ColumnSet("hil_warrantystatus", "hil_modelcode", "hil_modelname", "hil_serialnumber", "msdyn_customerasset", "hil_observation", "msdyn_incidenttype");
        query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, JobId);
        return [.. service.RetrieveMultiple(query).Entities];
    }
    internal List<Entity> FetchJobProductsData(Guid incidentId)
    {
        query = new QueryExpression("msdyn_workorderproduct");
        query.ColumnSet = new ColumnSet("hil_replacedpart", "hil_part", "hil_quantity", "hil_totalamount");
        query.Criteria.AddCondition("msdyn_workorderincident", ConditionOperator.Equal, incidentId);
        return [.. service.RetrieveMultiple(query).Entities];
    }
    internal List<Entity> FetchJobServiceData(Guid incidentId)
    {
        query = new QueryExpression("msdyn_workorderservice");
        query.ColumnSet = new ColumnSet("msdyn_service");
        query.Criteria.AddCondition("msdyn_workorderincident", ConditionOperator.Equal, incidentId);
        return [.. service.RetrieveMultiple(query).Entities];
    }
    internal List<Entity> FetchJobs()
    {
        query = new QueryExpression("msdyn_workorder");
        query.ColumnSet = new ColumnSet("msdyn_name", "hil_productcategory", "hil_productsubcategory", "hil_customername", "hil_mobilenumber", "hil_pincode", "hil_alternate", "hil_warrantystatus", "hil_countryclassification", "hil_closureremarks", "hil_customercomplaintdescription", "hil_actualcharges", "msdyn_substatus");
        FilterExpression filter = new FilterExpression(LogicalOperator.Or);
        filter.AddCondition("substatus", ConditionOperator.Equal, new Guid("1727fa6c-fa0f-e911-a94e-000d3af060a1"));// Closed
        filter.AddCondition("substatus", ConditionOperator.Equal, new Guid("1527fa6c-fa0f-e911-a94e-000d3af060a1"));// Cancelled
        filter.AddCondition("substatus", ConditionOperator.Equal, new Guid("6c8f2123-5106-ea11-a811-000d3af057dd")); // KKG Audit Failed 
        return [.. service.RetrieveMultiple(query).Entities];
    }

}
