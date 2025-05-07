internal class Jobs(List<JobIncident> jobIncidents)
{
    public string JobId = string.Empty;
    public string ProductCategory = string.Empty;
    public string ProductSubCategory = string.Empty;
    public string CustomerName = string.Empty;
    public string PhoneNubmer = string.Empty;
    public string PINCode = string.Empty;
    public string AlternateNumber = string.Empty;
    public string WarrantyStatus = string.Empty;
    public string CountryClassification = string.Empty;
    public string ClosureRemarks = string.Empty;
    public string CustomerComplaintDescription = string.Empty;
    public string ActualCharges = string.Empty;
    public string SubStatus = string.Empty;

    public List<JobIncident> Incidents = jobIncidents;
}
internal class JobIncident(List<JobProduct> jobProducts, List<JobService> jobServices)
{
    public string WarrantyStatus = string.Empty;
    public string ModelCode = string.Empty;
    public string ModelName = string.Empty;
    public string SerialNumber = string.Empty;
    public string InvoiceNumber = string.Empty;
    public string InvoiceDate = string.Empty;
    public string Observation = string.Empty;
    public string Cause = string.Empty;
    public List<JobProduct> Products = jobProducts;
    public List<JobService> Services = jobServices;
}
internal class JobProduct
{
    public string SparePartCode;
    public string SparePartDescription;
    public string Quantity;
    public string Amount;

    public JobProduct()
    {
        this.SparePartCode = string.Empty;
        this.SparePartDescription = string.Empty;
        this.Quantity = string.Empty;
        this.Amount = string.Empty;
    }
}
internal class JobService
{
    public string Action; 
    public JobService()
    {
        this.Action = string.Empty;
    }
}