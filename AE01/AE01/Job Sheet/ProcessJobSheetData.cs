using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

internal class ProcessJobSheetData(IOrganizationService _service)
{
    private readonly IOrganizationService service = _service;
    private readonly FetchJobsSheetData FetchData = new FetchJobsSheetData(_service);

    internal void ProcessIncidentData(Entity a, ref List<JobIncident> b)
    {
        JobIncident t = new JobIncident(new List<JobProduct>(), new List<JobService>()); 
        t.WarrantyStatus = a.Contains("hil_warrantystatus") ? a.FormattedValues["hil_warrantystatus"] : string.Empty;
        t.ModelCode = a.Contains("hil_modelcode") ? a.GetAttributeValue<EntityReference>("hil_modelcode").Name : string.Empty;
        t.ModelName = a.Contains("hil_modelname") ? a.GetAttributeValue<string>("hil_modelname") : string.Empty;
        t.SerialNumber = a.Contains("hil_serialnumber") ? a.GetAttributeValue<string>("hil_serialnumber") : string.Empty;
        t.Observation = a.Contains("hil_observation") ? a.GetAttributeValue<EntityReference>("hil_observation").Name : string.Empty;
        t.Cause = a.Contains("msdyn_incidenttype") ? a.GetAttributeValue<EntityReference>("msdyn_incidenttype").Name : string.Empty;
        EntityReference customerAssetRef = a.Contains("msdyn_customerasset") ? a.GetAttributeValue<EntityReference>("msdyn_customerasset") : new EntityReference();
        Entity c = service.Retrieve("msdyn_customerasset", customerAssetRef.Id, new ColumnSet("hil_invoiceno", "hil_invoicedate"));
        t.InvoiceDate = c.Contains("hil_invoiceno") ? c.GetAttributeValue<string>("hil_invoiceno") : string.Empty;
        t.InvoiceDate = c.Contains("hil_invoicedate") ? c.GetAttributeValue<DateTime>("hil_invoicedate").ToString() : string.Empty;
        b.Add(t);

        List<Entity> productsData = FetchData.FetchJobProductsData(a.Id);
        for (int i = 0; i < productsData.Count; i++)
        {
            ProcessProductsData(productsData[i], ref t.Products);
        }

        List<Entity> serviceData = FetchData.FetchJobServiceData(a.Id);
        for (int i = 0; i < serviceData.Count; i++)
        {
            ProcessServiceData(serviceData[i], ref t.Services);
        }

    }
    internal void ProcessProductsData(Entity a, ref List<JobProduct> b)
    {
        JobProduct t = new JobProduct();
        t.SparePartCode = a.Contains("hil_replacedpart") ? a.GetAttributeValue<EntityReference>("hil_replacedpart").Name : string.Empty;
        t.SparePartDescription = a.Contains("hil_part") ? a.GetAttributeValue<string>("hil_part") : string.Empty;
        t.Quantity = a.Contains("hil_quantity") ? a.GetAttributeValue<decimal>("hil_quantity").ToString() : string.Empty;
        t.Amount = a.Contains("hil_totalamount") ? a.GetAttributeValue<decimal>("hil_totalamount").ToString() : string.Empty;
        b.Add(t);
    }
    internal void ProcessServiceData(Entity a, ref List<JobService> b)
    {
        JobService t = new JobService();
        t.Action = a.Contains("msdyn_service") ? a.GetAttributeValue<EntityReference>("msdyn_service").Name : string.Empty;
        b.Add(t);
    }
    internal void ProcessJobData(ref List<Jobs> a, Entity j)
    {
        Jobs t = new Jobs(new List<JobIncident>());
        t.JobId = j.Contains("msdyn_name") ? j.GetAttributeValue<string>("msdyn_name") : string.Empty;
        t.ProductCategory = j.Contains("hil_productcategory") ? j.GetAttributeValue<EntityReference>("hil_productcategory").Name : string.Empty;
        t.ProductSubCategory = j.Contains("hil_productsubcategory") ? j.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : string.Empty;
        t.CustomerName = j.Contains("hil_customername") ? j.GetAttributeValue<string>("hil_customername") : string.Empty;
        t.PhoneNubmer = j.Contains("hil_mobilenumber") ? j.GetAttributeValue<string>("hil_mobilenumber") : string.Empty;
        t.PINCode = j.Contains("hil_pincode") ? j.GetAttributeValue<EntityReference>("hil_pincode").Name : string.Empty;
        t.AlternateNumber = j.Contains("hil_alternate") ? j.GetAttributeValue<string>("hil_alternate") : string.Empty;
        t.WarrantyStatus = j.Contains("hil_warrantystatus") ? j.FormattedValues["hil_warrantystatus"].ToString() : string.Empty;
        t.CountryClassification = j.Contains("hil_countryclassification") ? j.FormattedValues["hil_countryclassification"].ToString() : string.Empty;
        t.ClosureRemarks = j.Contains("hil_closureremarks") ? j.GetAttributeValue<string>("hil_closureremarks") : string.Empty;
        t.CustomerComplaintDescription = j.Contains("hil_customercomplaintdescription") ? j.GetAttributeValue<string>("hil_customercomplaintdescription") : string.Empty;
        t.ActualCharges = j.Contains("hil_actualcharges") ? j.GetAttributeValue<Money>("hil_actualcharges").Value.ToString() ?? string.Empty : string.Empty;
        t.SubStatus = j.Contains("msdyn_substatus") ? j.GetAttributeValue<EntityReference>("msdyn_substatus").Name : string.Empty;
        a.Add(t);
        List<Entity> jobIncidentData = FetchData.FetchJobIncidentData(j.Id);
        for (int i = 0; i < jobIncidentData.Count; i++)
        {
            ProcessIncidentData(jobIncidentData[i], ref t.Incidents);
        }
    }

}
