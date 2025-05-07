using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class UploadDocumentToAzureBlob : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string[] Source = { "1", "3", "5", "6", "7", "9", "12" };
            string JsonResponse = string.Empty;
            string DocumentName = null;
            string FileNumber = null;
            try
            {
                DocumentName = Convert.ToString(context.InputParameters["DocumentName"]);
                if (string.IsNullOrWhiteSpace(DocumentName))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "DocumentName is required"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                string pattern = @"^[^.\\]+(\.[^.\\]+)+$";
                bool isValidFileExtension = Regex.IsMatch(DocumentName, pattern);
                if (DocumentName.Length > 50)
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "Maximum length for DocumentName is 50"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (!isValidFileExtension)
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = ".Extension is Required for DocumentName field"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                string DocumentTypeGuid = Convert.ToString(context.InputParameters["DocumentType"]);
                if (string.IsNullOrWhiteSpace(DocumentTypeGuid))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "DocumentType is required"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else if (!APValidate.IsvalidGuid(DocumentTypeGuid))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "Invalid DocumentTypeGuid."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else if (DocumentTypeGuid.Length > 40)
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "Maximum length for DocumentType is 40"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                Guid MediaTypeGuid = Guid.Parse(DocumentTypeGuid);

                string DocumentBase64 = Convert.ToString(context.InputParameters["DocumentBase64"]);
                if (string.IsNullOrWhiteSpace(DocumentBase64))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "DocumentBase64 is required"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (!IsValidFile(DocumentName, DocumentBase64))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "Invalid file type. Only images and PDFs are allowed."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                string SourceType = Convert.ToString(context.InputParameters["SourceType"]);
                if (string.IsNullOrWhiteSpace(SourceType))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "Source Type required."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (!Source.Contains(SourceType))
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "Invalid SourceType."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                string RegardingGuid = Convert.ToString(context.InputParameters["RegardingGuid"]);
                if (!APValidate.IsvalidGuid(RegardingGuid))
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_workorderid'/>
                        <filter type='and'>
                            <condition attribute='msdyn_name' operator='eq' value='{RegardingGuid}'/> 
                        </filter>
                        </entity>
                        </fetch>";

                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        RegardingGuid = Convert.ToString(entCol.Entities[0].GetAttributeValue<Guid>("msdyn_workorderid"));
                    }
                }
                Entity entMediaType = service.Retrieve("hil_mediatype", new Guid(DocumentTypeGuid), new ColumnSet("hil_name"));
                string MediaTypeName = entMediaType.Contains("hil_name") ? entMediaType.GetAttributeValue<string>("hil_name") : string.Empty;

                if (string.IsNullOrWhiteSpace(RegardingGuid))
                {
                    FileNumber = Guid.NewGuid().ToString() + "_" + MediaTypeName;
                }
                else
                {
                    if (!APValidate.IsvalidGuid(RegardingGuid))
                    {
                        JsonResponse = JsonSerializer.Serialize(new Productresponse
                        {
                            StatusCode = "204",
                            Message = "Invalid RegardingGuid."
                        });
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else
                    {

                        FileNumber = RegardingGuid + "_" + MediaTypeName;
                    }
                }
                string RegardingType = Convert.ToString(context.InputParameters["RegardingType"]);
                string URL = StoreFileInAzureBlob(DocumentBase64, DocumentName, FileNumber);

                if (!string.IsNullOrWhiteSpace(URL) && RegardingGuid != null && APValidate.IsvalidGuid(RegardingGuid))
                {
                    JsonResponse = JsonSerializer.Serialize(ProductUpload(service, URL, FileNumber, new Guid(DocumentTypeGuid), RegardingType, RegardingGuid));
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else if (URL != null)
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        URL = URL,
                        StatusCode = "200",
                        Message = "ok"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else
                {
                    JsonResponse = JsonSerializer.Serialize(new Productresponse
                    {
                        StatusCode = "204",
                        Message = "URL not getting from server"
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                var retObj = JsonSerializer.Serialize(new Productresponse { StatusCode = "500", Message = "D365 Internal Server Error : " + ex.Message.ToUpper() });
                context.OutputParameters["data"] = retObj;
                return;
            }
        }
        public static string StoreFileInAzureBlob(string DocumentBase64, string DocumentName, string FileNumber)
        {
            const string connectionString = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
            string extension = Path.GetExtension(DocumentName);
            string containerName = "images";
            string blobName = FileNumber + extension;
            // Parse the connection string and get a reference to the storage account
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference(containerName);
            // Create the container if it does not exist
            container.CreateIfNotExists();
            // Get a reference to the blob
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(blobName);
            // Convert base64 to byte array and upload
            byte[] fileBytes = Convert.FromBase64String(DocumentBase64);
            using (MemoryStream stream = new MemoryStream(fileBytes))
            {
                blockBlob.UploadFromStream(stream);
            }
            return blockBlob.Uri.ToString();
        }
        public Productresponse ProductUpload(IOrganizationService service, string URL, string FileNumber, Guid DocumentTypeGuid, string RegardingType, string RegardingGuid)
        {
            Productresponse _Productresponse = new Productresponse();
            try
            {
                var MediaGallery = new Entity("hil_mediagallery");
                MediaGallery["hil_name"] = FileNumber;
                MediaGallery["hil_mediatype"] = new EntityReference("hil_mediatype", DocumentTypeGuid);
                MediaGallery["hil_url"] = URL;

                if (RegardingType.ToLower() == "msdyn_customerasset")
                {
                    Entity CustomerAssetInfo = service.Retrieve("msdyn_customerasset", new Guid(RegardingGuid), new ColumnSet("msdyn_name", "hil_customer"));
                    EntityReference Customer = CustomerAssetInfo.GetAttributeValue<EntityReference>("hil_customer");
                    MediaGallery["hil_customerasset"] = new EntityReference("msdyn_customerasset", CustomerAssetInfo.Id);
                    MediaGallery["hil_consumer"] = new EntityReference("contact", Customer.Id);
                }
                else if (RegardingType.ToLower() == "msdyn_workorder")
                {
                    Entity JobInfo = service.Retrieve("msdyn_workorder", new Guid(RegardingGuid), new ColumnSet("hil_customerref", "msdyn_name"));
                    EntityReference Customer = JobInfo.GetAttributeValue<EntityReference>("hil_customerref");
                    MediaGallery["hil_job"] = new EntityReference("msdyn_customerasset", new Guid(RegardingGuid));
                    MediaGallery["hil_consumer"] = new EntityReference("contact", Customer.Id);
                }
                else if (RegardingType.ToLower() == "salesorder")
                {
                    Entity JobInfo = service.Retrieve("salesorder", new Guid(RegardingGuid), new ColumnSet("name", "customerid"));
                    EntityReference Customer = JobInfo.GetAttributeValue<EntityReference>("customerid");
                    MediaGallery["hil_salesorder"] = new EntityReference("salesorder", new Guid(RegardingGuid));
                    MediaGallery["hil_consumer"] = new EntityReference("contact", Customer.Id);
                }
                else
                {
                    _Productresponse.MediaGalleryGUID = null;
                    _Productresponse.URL = null;
                    _Productresponse.StatusCode = "200";
                    _Productresponse.Message = "Please enter valid RegardingType";
                    return _Productresponse;
                }
                Guid mediaGalleryGuid = service.Create(MediaGallery);
                _Productresponse.MediaGalleryGUID = mediaGalleryGuid.ToString();
                _Productresponse.URL = URL;
                _Productresponse.StatusCode = "200";
                _Productresponse.Message = "ok";

            }
            catch (Exception ex)
            {
                _Productresponse.StatusCode = "500";
                _Productresponse.Message = "D365 Internal Server Error : " + ex.Message.ToUpper();
            }
            return _Productresponse;
        }
        public static bool IsValidFile(string fileName, string base64Content)
        {
            string extension = Path.GetExtension(fileName).ToLower();
            string[] validExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".pdf" };
            if (!validExtensions.Contains(extension))
            {
                return false;
            }

            string mimeType = GetMimeType(base64Content);
            string[] validMimeTypes = { "image/jpeg", "image/png", "image/gif", "image/bmp", "application/pdf" };
            return validMimeTypes.Contains(mimeType);
        }
        public static string GetMimeType(string base64Content)
        {
            byte[] data = Convert.FromBase64String(base64Content);
            using (MemoryStream ms = new MemoryStream(data))
            {
                byte[] buffer = new byte[256];
                ms.Read(buffer, 0, 256);
                string hex = BitConverter.ToString(buffer).Replace("-", string.Empty).ToLower();

                if (hex.StartsWith("ffd8ff"))
                {
                    return "image/jpeg";
                }
                if (hex.StartsWith("89504e47"))
                {
                    return "image/png";
                }
                if (hex.StartsWith("47494638"))
                {
                    return "image/gif";
                }
                if (hex.StartsWith("424d"))
                {
                    return "image/bmp";
                }
                if (hex.StartsWith("25504446"))
                {
                    return "application/pdf";
                }
            }
            return string.Empty;
        }
    }
    public class Productresponse
    {
        public string StatusCode { get; set; }
        public string Message { get; set; }
        public string URL { get; set; }
        public string MediaGalleryGUID { get; set; }

    }
}
