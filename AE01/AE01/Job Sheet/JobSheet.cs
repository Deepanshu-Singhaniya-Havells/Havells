using Azure.Storage.Blobs;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;


class JobSheet(IOrganizationService _service)
{
    private readonly FetchJobsSheetData FetchData = new FetchJobsSheetData(_service);
    private readonly ProcessJobSheetData ProcessData = new ProcessJobSheetData(_service);
    private readonly AzureClient AzureClient = new AzureClient();
       
    internal async Task FetchJSON()
    {
        try
        {
            List<Jobs> toReturn = new List<Jobs>();
            List<Entity> jobsList = FetchData.FetchJobs();
            for (int i = 0; i < 10; i++)
            {
                ProcessData.ProcessJobData(ref toReturn, jobsList[i]);
            }
            string dataToUpload = JsonConvert.SerializeObject(toReturn);
            BlobContainerClient blobClient = AzureClient.CreateConnection("jobsheet-json");
            await AzureClient.UploadFromStringAsync(blobClient, "First Data Upload", dataToUpload);
        }
        catch (Exception ex)
        {
            throw new Exception(ex.ToString());
        }
    }
}
