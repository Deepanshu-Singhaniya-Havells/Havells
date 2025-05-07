using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Grievance_Escalation
{
    public class CreateCase : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            try
            {
                string CustomerGuid = Convert.ToString(context.InputParameters["CustomerGuid"]);
                string AddressGuid = Convert.ToString(context.InputParameters["AddressGuid"]);
                string ServiceRequestGuid = Convert.ToString(context.InputParameters["ServiceRequestGuid"]);
                string ComplaintCategoryId = Convert.ToString(context.InputParameters["ComplaintCategoryId"]);
                string ComplaintTitle = Convert.ToString(context.InputParameters["ComplaintTitle"]);
                string ComplaintDescription = Convert.ToString(context.InputParameters["ComplaintDescription"]);
                string Attachment = Convert.ToString(context.InputParameters["Attachment"]);
                string FileName = Convert.ToString(context.InputParameters["FileName"]);

                if (string.IsNullOrWhiteSpace(AddressGuid))
                {
                    var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "AddressGuid cannot be empty" });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;

                }
                if (!string.IsNullOrWhiteSpace(AddressGuid))
                {
                    if (!APValidate.IsvalidGuid(AddressGuid))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "AddressGuid is not Valid" });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }
                }

                if (!string.IsNullOrWhiteSpace(ServiceRequestGuid))
                {
                    if (!APValidate.IsvalidGuid(ServiceRequestGuid))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "ServiceRequestGuid is not Valid" });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }
                }
                if (string.IsNullOrWhiteSpace(ComplaintCategoryId))
                {
                    var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "ComplaintCategoryId cannot be empty" });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;
                }
                if (!string.IsNullOrWhiteSpace(ComplaintCategoryId))
                {
                    if (!APValidate.IsvalidGuid(ComplaintCategoryId))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "ComplaintCategoryId is not Valid" });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }
                }
                if (string.IsNullOrWhiteSpace(ComplaintTitle))
                {
                    var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "ComplaintTitle cannot be empty" });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;
                }
                if (string.IsNullOrWhiteSpace(CustomerGuid))
                {

                    var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "CustomerGuid required" });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;

                }
                if (!string.IsNullOrWhiteSpace(CustomerGuid))
                {
                    if (!APValidate.IsvalidGuid(CustomerGuid))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "CustomerGuid is not Valid" });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(FileName))
                {
                    string[] allowedExtensions = { ".png", ".jpg", ".jpeg", ".pdf" };
                    string fileExtension = (Path.GetExtension(FileName) ?? "").ToLower();

                    if (string.IsNullOrWhiteSpace(fileExtension) || !allowedExtensions.Contains(fileExtension))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "Invalid/missing file extension. Allowed extensions are: png, jpg, jpeg, pdf." });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(FileName) || !string.IsNullOrWhiteSpace(Attachment))
                {
                    if (string.IsNullOrWhiteSpace(FileName))
                    {

                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "FileName is required when Attachment is provided." });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(Attachment))
                    {
                        var jobIdResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.BadRequest, Message = "Attachment is required when FileName is provided." });
                        context.OutputParameters["data"] = jobIdResponse;
                        return;
                    }
                }
                var obj = CreateCaserecord(service, new CreateCase_Request
                {
                    CustomerGuid = CustomerGuid,
                    AddressGuid = AddressGuid,
                    ServiceRequestGuid = ServiceRequestGuid,
                    ComplaintCategoryId = ComplaintCategoryId,
                    ComplaintTitle = ComplaintTitle,
                    ComplaintDescription = ComplaintDescription,
                    Attachment = Attachment,
                    FileName = FileName
                });
                if (obj.Item2.StatusCode == (int)HttpStatusCode.OK)
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(obj.Item1);
                }
                else
                {
                    var RequestStatus = new
                    {
                        StatusCode = obj.Item2.StatusCode,
                        Message = obj.Item2.Message
                    };
                    context.OutputParameters["data"] = JsonSerializer.Serialize(RequestStatus);
                }
            }
            catch (Exception ex)
            {
                var TechnicianScoreValidResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = (int)HttpStatusCode.InternalServerError, Message = ex.Message });
                context.OutputParameters["data"] = TechnicianScoreValidResponse;
                return;
            }
        }
        public (CreateCase_Response, RequestStatus) CreateCaserecord(IOrganizationService _CrmService, CreateCase_Request obj)
        {
            CreateCase_Response res = new CreateCase_Response();
            res.ServiceRequests = new List<CreateCase_ServiceRequest>();
            try
            {
                OptionSetValue CaseType;
                EntityReference CaseDepartment;
                EntityReference Branch;
                EntityReference Product;

                if (obj.ComplaintCategoryId != null)
                {
                    Entity caseCategory = _CrmService.Retrieve("hil_casecategory", new Guid(obj.ComplaintCategoryId), new ColumnSet("hil_casedepartment", "hil_casetype"));
                    CaseType = caseCategory.GetAttributeValue<OptionSetValue>("hil_casetype");
                    CaseDepartment = caseCategory.GetAttributeValue<EntityReference>("hil_casedepartment");
                }
                else
                {

                    return (res, new RequestStatus());
                }
                if (string.IsNullOrWhiteSpace(obj.ServiceRequestGuid) && CaseDepartment.Id.ToString() == "ab3dbc3d-4e6e-ee11-8179-6045bdac526a") // service
                {
                    return (res, new RequestStatus
                    {
                        StatusCode = (int)HttpStatusCode.NotFound,
                        Message = "ServiceRequestID is Required when Case Department is 'Service'."
                    });
                }
                if (!string.IsNullOrWhiteSpace(obj.ServiceRequestGuid))
                {
                    Entity Jobs = _CrmService.Retrieve("msdyn_workorder", new Guid(obj.ServiceRequestGuid), new ColumnSet("hil_branch", "hil_productcategory"));
                    Branch = Jobs.GetAttributeValue<EntityReference>("hil_branch");
                    Product = Jobs.GetAttributeValue<EntityReference>("hil_productcategory");
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(obj.AddressGuid))
                    {
                        Entity Address = _CrmService.Retrieve("hil_address", new Guid(obj.AddressGuid), new ColumnSet("hil_branch"));
                        Branch = Address.GetAttributeValue<EntityReference>("hil_branch");
                        Product = null;
                    }
                    else
                    {

                        return (res, new RequestStatus());
                    }
                }
                var complaintEntity = new Entity("incident");
                complaintEntity["customerid"] = new EntityReference("contact", new Guid(obj.CustomerGuid));
                complaintEntity["hil_address"] = new EntityReference("hil_address", new Guid(obj.AddressGuid));
                if (!string.IsNullOrWhiteSpace(obj.ServiceRequestGuid))
                {
                    complaintEntity["hil_job"] = new EntityReference("msdyn_workorder", new Guid(obj.ServiceRequestGuid));
                }
                complaintEntity["hil_casecategory"] = new EntityReference("hil_casecategory", new Guid(obj.ComplaintCategoryId));
                complaintEntity["title"] = obj.ComplaintTitle;
                complaintEntity["description"] = obj.ComplaintDescription;
                complaintEntity["caseorigincode"] = new OptionSetValue(6); // Havells One Website
                complaintEntity["casetypecode"] = new OptionSetValue(CaseType.Value);
                complaintEntity["productid"] = Product;
                complaintEntity["hil_branch"] = Branch;
                complaintEntity["hil_casedepartment"] = CaseDepartment;
                Guid complaintId = _CrmService.Create(complaintEntity);

                Entity Case = _CrmService.Retrieve("incident", complaintId, new ColumnSet("ticketnumber"));
                string ComplaintGuid = Case.GetAttributeValue<string>("ticketnumber");

                string url = "";
                if (!string.IsNullOrWhiteSpace(obj.Attachment) && !string.IsNullOrWhiteSpace(obj.FileName))
                {
                    url = StoreFileInAzureBlob(obj.Attachment, obj.FileName, complaintId.ToString());

                    if (!string.IsNullOrWhiteSpace(url))
                    {
                        var MediaGallery = new Entity("hil_mediagallery");
                        MediaGallery["hil_name"] = complaintId.ToString() + "_Attachment";
                        MediaGallery["hil_mediatype"] = new EntityReference("hil_mediatype", new Guid("f995d6d0-968f-ef11-8a6a-6045bdaa9bfe")); // Grievance Attachment
                        MediaGallery["hil_consumer"] = new EntityReference("contact", new Guid(obj.CustomerGuid));
                        MediaGallery["hil_url"] = url;
                        if (!string.IsNullOrWhiteSpace(obj.ServiceRequestGuid))
                        {
                            MediaGallery["hil_job"] = new EntityReference("contact", new Guid(obj.ServiceRequestGuid));
                        }
                        MediaGallery["hil_case"] = new EntityReference("contact", complaintId);

                        Guid mediaGalleryGuid = _CrmService.Create(MediaGallery);

                    }
                }
                CreateCase_ServiceRequest tempComplaint = new CreateCase_ServiceRequest();
                tempComplaint.ComplaintId = ComplaintGuid;
                tempComplaint.ComplaintGuid = complaintId.ToString();
                tempComplaint.Attachment_URL = url;
                res.ServiceRequests.Add(tempComplaint);

                return (res, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {

                return (res, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error :" + ex.Message.ToUpper()
                });
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
    }
    public class CreateCase_ServiceRequest
    {
        public string ComplaintId { get; set; }
        public string ComplaintGuid { get; set; }
        public string Attachment_URL { get; set; }
    }
    public class CreateCase_Response
    {
        public List<CreateCase_ServiceRequest> ServiceRequests { get; set; }
    }
    public class CreateCase_Request
    {
        public string CustomerGuid { get; set; }
        public string AddressGuid { get; set; }
        public string ServiceRequestGuid { get; set; }
        public string ComplaintCategoryId { get; set; }
        public string ComplaintTitle { get; set; }
        public string ComplaintDescription { get; set; }
        public string Attachment { get; set; }
        public string FileName { get; set; }

    }
}
