using Azure.Storage.Blobs;

class AzureClient
{
    internal BlobContainerClient CreateConnection(string containerName)
    {
        // Retrieve the connection string for use with the application. 
        string connectionString = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
        // Create a BlobServiceClient object 
        var blobServiceClient = new BlobServiceClient(connectionString);

        var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
        //container.UploadBlob()
        Console.WriteLine("Connection Created Successfully");
        return containerClient;
    }

    internal async Task UploadFromStringAsync(BlobContainerClient containerClient, string blobName, string blobContents)
    {
        try
        {
            BlobClient blobClient = containerClient.GetBlobClient(blobName);
            await blobClient.UploadAsync(BinaryData.FromString(blobContents), overwrite: true);
            Console.WriteLine("Uploaded to blob");
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

}

