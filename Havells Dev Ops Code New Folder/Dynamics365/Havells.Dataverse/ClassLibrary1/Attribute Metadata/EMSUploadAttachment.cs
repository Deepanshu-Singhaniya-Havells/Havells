using Microsoft.Crm.Sdk.Messages;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Attribute_Metadata
{
    public class EMSUploadAttachment : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            ResponseUpload responseUpload;
            #region Value extract
            UploadAttachment reqParm = new UploadAttachment
            {
                RegardingGuid = Convert.ToString(context.InputParameters["RegardingGuid"]),
                Department = Convert.ToString(context.InputParameters["Department"]),
                DocGuid = Convert.ToString(context.InputParameters["DocGuid"]),
                DocumentType = Convert.ToString(context.InputParameters["DocumentType"]),
                FileName = Convert.ToString(context.InputParameters["FileName"]),
                FileSize = Convert.ToString(context.InputParameters["FileSize"]),
                FileString = Convert.ToString(context.InputParameters["FileString"]),
                ObjectType = Convert.ToString(context.InputParameters["ObjectType"]),
                Source = Convert.ToString(context.InputParameters["Source"]),
                Subject = Convert.ToString(context.InputParameters["Subject"]),
                ValidFrom = Convert.ToString(context.InputParameters["ValidFrom"]),
                ValidTill = Convert.ToString(context.InputParameters["ValidTill"]),
                IsDeleted = Convert.ToBoolean(context.InputParameters["IsDeleted"])
            };
            #endregion Value extract
            #region Validation Check
            if (context.InputParameters.Contains("RegardingGuid") && !string.IsNullOrWhiteSpace(reqParm.RegardingGuid))
            {
                if (!APValidate.IsvalidGuid(reqParm.RegardingGuid))
                {
                    responseUpload = new ResponseUpload() { Message = "RegardingGuid is not valid." };
                    context.OutputParameters["data"] = JsonSerializer.Serialize(responseUpload);
                    return;
                }
            }
            if (context.InputParameters.Contains("Department") && !string.IsNullOrWhiteSpace(reqParm.Department))
            {
                if (!APValidate.IsValidString(reqParm.Department))
                {
                    responseUpload = new ResponseUpload() { Message = "Department is not valid." };
                    context.OutputParameters["data"] = JsonSerializer.Serialize(responseUpload);
                    return;
                }
            }
            if (context.InputParameters.Contains("DocGuid") && !string.IsNullOrWhiteSpace(reqParm.DocGuid))
            {
                if (!APValidate.IsvalidGuid(reqParm.DocGuid))
                {
                    responseUpload = new ResponseUpload() { Message = "DocGuid is not valid." };
                    context.OutputParameters["data"] = JsonSerializer.Serialize(responseUpload);
                    return;
                }
            }
            #endregion
            ResponseUpload _response = UploadAttachment(reqParm, service);
            context.OutputParameters["data"] = JsonSerializer.Serialize(_response);
        }
        public ResponseUpload UploadAttachment(UploadAttachment reqParm, IOrganizationService service)
        {
            ResponseUpload resParm = new ResponseUpload();
            String blobURL = string.Empty;
            if (reqParm.DocGuid != string.Empty && reqParm.DocGuid != null)
            {
                if (!reqParm.IsDeleted)
                {
                    resParm.Message = "Parameter Mismatch to Delete Existing Document";
                    resParm.Status = false;
                    return resParm;
                }
                else
                {
                    Entity ent = new Entity("hil_attachment", new Guid(reqParm.DocGuid));
                    ent["hil_isdeleted"] = true;
                    service.Update(ent);
                    resParm.Message = "Success";
                    resParm.Status = true;
                    return resParm;
                }
            }
            else
            {
                if (reqParm.FileName == null || reqParm.FileName == "" ||
                reqParm.FileSize == null || reqParm.FileSize == "" ||
                reqParm.FileString == null || reqParm.FileString == "" ||
                reqParm.DocumentType == null || reqParm.DocumentType == "")
                {
                    resParm.Message = "Some Required feild is missing";
                    resParm.Status = false;
                    return resParm;
                }
                else
                {
                    try
                    {
                        EntityReference _erDepartment = null;
                        QueryExpression query = null;
                        EntityCollection entCol = null;

                        EntityReference _doctypee = new EntityReference(); //getOptionSetValue(reqParm.DocumentType, "hil_attachment", "hil_doctype", service);
                        Guid loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
                        query = new QueryExpression("hil_attachmentdocumenttype");
                        query.ColumnSet = new ColumnSet("hil_containername", "hil_department");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, reqParm.DocumentType.Trim());

                        entCol = service.RetrieveMultiple(query);
                        String containerName = string.Empty;

                        if (entCol.Entities.Count == 0)
                        {
                            resParm.Message = "Given Document Type is not Defined in Dynamics.";
                            resParm.Status = false;
                            return resParm;
                        }
                        else
                        {
                            _doctypee = new EntityReference(entCol[0].LogicalName, entCol[0].Id);
                            containerName = entCol[0].Contains("hil_containername") ? entCol[0].GetAttributeValue<string>("hil_containername") : "devanduat";
                            if (entCol.Entities[0].Contains("hil_department"))
                                _erDepartment = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_department");
                        }
                        string[] ext = reqParm.FileName.Split('.');
                        String fileName = string.Empty;

                        if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
                        {
                            if (reqParm.ObjectType == "hil_tender")
                            {
                                fileName = service.Retrieve(reqParm.ObjectType, new Guid(reqParm.RegardingGuid), new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
                            }
                            else
                            {
                                fileName = reqParm.RegardingGuid;
                            }
                            fileName = fileName + "_" + reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                        }
                        else
                        {
                            fileName = reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                        }
                        //String fileName = reqParm.RegardingGuid + "_" + reqParm.DocumentType + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + "." + ext[ext.Length - 1];
                        byte[] fileContent = Convert.FromBase64String(reqParm.FileString);
                        try
                        {
                            blobURL = Upload(fileName, fileContent, containerName);  // reqParm.ContainerName); ;
                        }
                        catch (Exception ex)
                        {
                            resParm.Message = ex.Message;
                            resParm.Status = false;
                            return resParm;
                        }

                        Entity _attachment = new Entity("hil_attachment");
                        if (reqParm.Subject == string.Empty || reqParm.Subject.Trim().Length == 0)
                        {
                            int _rowCount = 0;
                            int pageNumber = 1;
                            query = new QueryExpression("hil_attachment");
                            query.ColumnSet = new ColumnSet(false);
                            query.PageInfo = new PagingInfo();
                            query.PageInfo.Count = 5000;
                            query.PageInfo.PageNumber = pageNumber;
                            query.PageInfo.PagingCookie = null;
                            while (true)
                            {
                                EntityCollection entColFile = service.RetrieveMultiple(query);
                                _rowCount += entColFile.Entities.Count;
                                if (entColFile.MoreRecords)
                                {
                                    // Increment the page number to retrieve the next page.
                                    query.PageInfo.PageNumber++;
                                    // Set the paging cookie to the paging cookie returned from current results.
                                    query.PageInfo.PagingCookie = entColFile.PagingCookie;
                                }
                                else
                                {
                                    // If no more records are in the result nodes, exit the loop.
                                    break;
                                }
                            }
                            _rowCount += 1;
                            _attachment["subject"] = "HIL_" + _rowCount.ToString().PadLeft(8, '0');
                            _attachment["description"] = reqParm.Description;
                        }
                        else
                        {
                            _attachment["subject"] = reqParm.Subject;
                            _attachment["description"] = reqParm.Description;
                        }

                        if (_erDepartment != null) { _attachment["hil_department"] = _erDepartment; }

                        _attachment["hil_documenttype"] = _doctypee;
                        _attachment["hil_docurl"] = blobURL;
                        resParm.BlobURL = blobURL;
                        _attachment["hil_docsize"] = double.Parse(reqParm.FileSize);
                        if (reqParm.ValidFrom != null && reqParm.ValidFrom != "")
                        {
                            DateTime fromDate = new DateTime(Convert.ToInt32(reqParm.ValidFrom.Substring(0, 4)),
                                Convert.ToInt32(reqParm.ValidFrom.Substring(4, 2)),
                                Convert.ToInt32(reqParm.ValidFrom.Substring(6, 2)));

                            _attachment["scheduledstart"] = fromDate.Date;
                        }
                        if (reqParm.ValidTill != null && reqParm.ValidTill != "")
                        {
                            DateTime toDate = new DateTime(Convert.ToInt32(reqParm.ValidTill.Substring(0, 4)),
                               Convert.ToInt32(reqParm.ValidTill.Substring(4, 2)),
                               Convert.ToInt32(reqParm.ValidTill.Substring(6, 2)));
                            _attachment["scheduledend"] = toDate.Date;
                        }

                        if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
                        {

                            if (reqParm.ObjectType != "systemuser")
                                _attachment["regardingobjectid"] = new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid));
                            else
                                _attachment["hil_regardinguser"] = new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid));
                        }

                        _attachment["hil_sourceofdocument"] = new OptionSetValue(int.Parse(reqParm.Source));
                        try
                        {
                            service.Create(_attachment);
                            if (reqParm.RegardingGuid != null && reqParm.RegardingGuid != string.Empty)
                            {
                                try
                                {
                                    Entity _entityReg = service.Retrieve(reqParm.ObjectType, new Guid(reqParm.RegardingGuid), new ColumnSet("ownerid"));
                                    if (_entityReg.Contains("ownerid"))
                                    {
                                        EntityReference owner = _entityReg.GetAttributeValue<EntityReference>("ownerid");
                                        if (loginUserGuid != owner.Id)
                                        {
                                            Assign(owner, new EntityReference(reqParm.ObjectType, new Guid(reqParm.RegardingGuid)), service);
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            resParm.Message = "Failed to Create Record : " + ex.Message;
                            resParm.Status = false;
                            return resParm;
                        }
                        resParm.Message = "File Uplaoded Sucessfully";
                        resParm.Status = true;
                        return resParm;
                    }
                    catch (Exception ex)
                    {
                        resParm.Message = "Failed to Create Record : " + ex.Message;
                        resParm.Status = false;
                        return resParm;
                    }
                }
            }
        }
        string Upload(string fileName, byte[] fileContent, string containerName)
        {
            string _blobURI = string.Empty;
            try
            {
                fileName = Regex.Replace(fileName, @"\s+", String.Empty);
                //byte[] fileContent = Convert.FromBase64String(noteBody);
                string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
                // create object of storage account
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConnectionSting);

                // create client of storage account
                CloudBlobClient client = storageAccount.CreateCloudBlobClient();

                // create the reference of your storage account
                CloudBlobContainer container = client.GetContainerReference(ToURLSlug(containerName));

                // check if the container exists or not in your account
                var isCreated = container.CreateIfNotExists();

                // set the permission to blob type
                container.SetPermissionsAsync(new BlobContainerPermissions
                { PublicAccess = BlobContainerPublicAccessType.Blob });

                // create the memory steam which will be uploaded
                using (MemoryStream memoryStream = new MemoryStream(fileContent))
                {
                    // set the memory stream position to starting
                    memoryStream.Position = 0;

                    // create the object of blob which will be created
                    // Test-log.txt is the name of the blob, pass your desired name
                    CloudBlockBlob blob = container.GetBlockBlobReference(fileName);

                    // get the mime type of the file
                    string mimeType = "application/unknown";
                    string ext = (fileName.Contains(".")) ?
                                System.IO.Path.GetExtension(fileName).ToLower() : "." + fileName;
                    Microsoft.Win32.RegistryKey regKey =
                                Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
                    if (regKey != null && regKey.GetValue("Content Type") != null)
                        mimeType = regKey.GetValue("Content Type").ToString();

                    // set the memory stream position to zero again
                    // this is one of the important stage, If you miss this step, 
                    // your blob will be created but size will be 0 byte
                    memoryStream.ToArray();
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    // set the mime type
                    blob.Properties.ContentType = mimeType;

                    // upload the stream to blob
                    blob.UploadFromStream(memoryStream);
                    _blobURI = blob.Uri.AbsoluteUri;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Failed to Upload File: " + ex.Message);
            }
            return _blobURI;
        }
        public void Assign(EntityReference _Assignee, EntityReference _Targetd, IOrganizationService service)
        {
            try
            {
                AssignRequest assign = new AssignRequest();
                assign.Assignee = _Assignee;
                assign.Target = _Targetd;
                service.Execute(assign);
            }
            catch
            {
            }
        }
        string ToURLSlug(string s)
        {
            return Regex.Replace(s, @"[^a-z0-9]+", "-", RegexOptions.IgnoreCase)
                .Trim(new char[] { '-' })
                .ToLower();
        }
    }
    public class ResponseUpload
    {
        public string Message { get; set; }
        public bool Status { get; set; }
        public string BlobURL { get; set; }
        public string DocGuid { get; set; }
    }
    public class UploadAttachment
    {

        public string Subject { get; set; }
        public string Description { get; set; }
        public string FileSize { get; set; }
        public string FileName { get; set; }
        public string FileString { get; set; }
        public string ObjectType { get; set; }
        public string RegardingGuid { get; set; }
        public string Source { get; set; }
        public string DocumentType { get; set; }
        public string ValidFrom { get; set; }
        public string ValidTill { get; set; }
        public bool IsDeleted { get; set; }
        public string DocGuid { get; set; }
        public string Department { get; set; }
    }
}

