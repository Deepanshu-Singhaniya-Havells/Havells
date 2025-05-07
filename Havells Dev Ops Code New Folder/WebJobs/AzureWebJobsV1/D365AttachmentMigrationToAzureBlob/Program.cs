using System;
using System.ServiceModel.Description;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;

namespace D365AttachmentMigrationToAzureBlob
{
    public class Program
    {
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _fromDate = string.Empty;
        static string _toDate = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havellscrmdev.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                ///Backup Start
                //CustomerAsset();
                //MigrateNotesToBlob();
                //Console.ReadKey();
                //WithOutRegarding();
                //Customer();
                //InventoryRequest();
                ///Backup End
                //DeleteBlobFile();
                //----------***------------
                //MigrateNotesToBlob();
                //Console.WriteLine("*************** Work Order ***************");
                //DownloadCustomerSignature();
                //Console.WriteLine("*************** Customer Asset ***************");
                //DownloadCustomerAssetImage();
                AuthenticateConsumer _ret= AuthenticateConsumerAMC(new AuthenticateConsumer() {JWTToken="",MobileNumber="8285906485",SourceOrigin="Havells SyncApp"});
            }
        }

        static AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam)
        {
            AuthenticateConsumer _retValue = new AuthenticateConsumer() { StatusCode = "200", StatusDescription = "OK", MobileNumber = requestParam.MobileNumber, SessionId = string.Empty, JWTToken = requestParam.JWTToken, SourceOrigin = requestParam.SourceOrigin };
            Entity _userSession = null;
            try
            {
                string _fetxhXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='contact'>
                <attribute name='contactid' />
                <filter type='and'>
                    <filter type='or'>
                        <condition attribute='emailaddress1' operator='eq' value='{requestParam.MobileNumber}' />
                        <condition attribute='mobilephone' operator='eq' value='{requestParam.MobileNumber}' />
                    </filter>
                </filter>
                </entity>
                </fetch>";
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetxhXML));
                if (entCol.Entities.Count > 0)
                {
                    QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                    queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                    ConditionExpression condExp = new ConditionExpression("hil_name", ConditionOperator.Equal, requestParam.MobileNumber);
                    queryExp.Criteria.AddCondition(condExp);
                    condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                    queryExp.Criteria.AddCondition(condExp);
                    queryExp.AddOrder("hil_expiredon", OrderType.Descending);

                    entCol = _service.RetrieveMultiple(queryExp);
                    if (entCol.Entities.Count > 0) //No Active session found
                    {
                        //DateTime _expiredOn = entCol.Entities[0].GetAttributeValue<DateTime>("hil_expiredon").AddMinutes(330);
                        //if (_expiredOn <= DateTime.Now)
                        //{
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = entCol.Entities[0].Id,
                                LogicalName = "hil_consumerloginsession",
                            },
                            State = new OptionSetValue(1), //Inactive
                            Status = new OptionSetValue(2) //Inactive
                        };
                        _service.Execute(setStateRequest);
                        //}
                    }
                    _userSession = new Entity("hil_consumerloginsession");
                    _userSession["hil_name"] = requestParam.MobileNumber;
                    _userSession["hil_origin"] = requestParam.SourceOrigin;
                    _userSession["hil_expiredon"] = DateTime.Now;
                    _retValue.SessionId = _service.Create(_userSession).ToString();
                }
                else {
                    _retValue.StatusCode = "204";
                    _retValue.StatusDescription = "Consumer Doesn't Exist in D365.";
                }
            }
            catch(Exception ex)
            {
                _retValue.StatusCode = "204";
                _retValue.StatusDescription = "ERROR !!! " + ex.Message;
            }
            return _retValue;
        }
        static void DeleteAuditData() {
            try
            {
                int i = 0;
                int j = 0;
                while (true)
                {
                    string fromDate = "01/01/2018";
                    string toDate = "12/31/2019";

                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='audit'>
                    <all-attributes />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='" + fromDate + @"' />
                        <condition attribute='createdon' operator='on-or-before' value='" + toDate + @"' />
                    </filter>
                    <link-entity name='hil_assignmentmatrix' from='hil_assignmentmatrixid' to='objectid' link-type='inner' alias='ae' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }

                    foreach (var record in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)record.Attributes["objectid"]);
                        var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                        j += 1;
                        DeleteAudit.Target = (EntityReference)record.Attributes["objectid"];
                        if (EntityLogReference.Id != Guid.Empty)
                        {
                            try
                            {
                                _service.Execute(DeleteAudit);
                            }
                            catch { }
                            i += 1;
                        }
                        Console.WriteLine(((DateTime)record.Attributes["createdon"]).ToString() + ":" + i.ToString() + "/" + j.ToString());
                    }
                }

            }
            catch (Exception err)
            {
                Console.WriteLine("{0} Exception caught.", err.Message);
            }
        }
        static void MigrateBlobToNotes()
        {
            string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
            CloudStorageAccount mycloudStorageAccount = CloudStorageAccount.Parse(ConnectionSting);
            CloudBlobClient blobClient = mycloudStorageAccount.CreateCloudBlobClient();

            Guid _currentRecId = Guid.Empty;
            int i = 0;
            int j = 0;
            string filename = string.Empty;
            string _blobURI = string.Empty;
            string documentbody = string.Empty;

            try
            {
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='annotation'>
                        <attribute name='objectid' />
                        <attribute name='subject' />
                        <attribute name='createdon' />
                        <attribute name='notetext' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                            <condition attribute='createdon' operator='this-month' />
                            <condition attribute='notetext' operator='like' value='%https://d365storagesa.blob%' />
                            <condition attribute='filesize' operator='eq' value='0' />
                        </filter>
                        </entity>
                        </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        try
                        {
                            String fileName = entity.GetAttributeValue<string>("notetext");
                            Uri blobUri = new Uri(fileName);
                            CloudBlockBlob blob = new CloudBlockBlob(blobUri, blobClient);
                            blob.FetchAttributes();//Fetch blob's properties
                            byte[] arr = new byte[blob.Properties.Length];
                            blob.DownloadToByteArray(arr, 0);
                            var azureBase64 = Convert.ToBase64String(arr);
                            j += 1;
                            byte[] NoteByte = Convert.FromBase64String(azureBase64);

                            entity["documentbody"] = Convert.ToBase64String(NoteByte);
                            entity["filename"] = entity.GetAttributeValue<string>("subject");
                            _service.Update(entity);
                            Console.WriteLine(i.ToString() + "/" + j.ToString());
                        }
                        catch {}
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ":FileName=" + filename);
            }
        }
        static void MigrateNotesToBlob()
        {
            Guid _currentRecId = Guid.Empty;
            int i = 1;
            int j = 0;
            string filename = string.Empty;
            DateTime createdon;
            int filesize;
            EntityReference objectid;
            bool isDocument;
            string _blobURI = string.Empty;
            Entity entTemp;
            string documentbody = string.Empty;
            string notetext = string.Empty;

            try
            {
                while (true)
                {
                    /* 
                    <condition attribute='createdon' operator='on-or-after' value='2019-02-01' />
                    <condition attribute='createdon' operator='on-or-before' value='2019-02-28' /> 
                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='objectid' link-type='inner' alias='ad' />
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='objectid' link-type='inner' alias='ae' />
                    <link-entity name='msdyn_workorderproduct' from='msdyn_workorderproductid' to='objectid' link-type='inner' alias='ai' />
                    <link-entity name='msdyn_workorderservice' from='msdyn_workorderserviceid' to='objectid' link-type='inner' alias='ai' />
                    <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='objectid' link-type='inner' alias='cj' />
                    <link-entity name='hil_inventoryrequest' from='hil_inventoryrequestid' to='objectid' link-type='inner' alias='bo' />
                    <link-entity name='msdyn_approval' from='activityid' to='objectid' link-type='inner' alias='ce' />
                    <link-entity name='hil_unitwarranty' from='hil_unitwarrantyid' to='objectid' link-type='inner' alias='cg' />
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='objectid' link-type='inner' alias='ae' >
                    <attribute name='msdyn_name' />
                    </link-entity>
                    <condition attribute='createdon' operator='on-or-after' value='2021-04-11' />
                    <condition attribute='createdon' operator='on-or-before' value='2021-04-20' /> 
                    <condition attribute='objectid' operator='not-null' />
                    <value>10396</value>
                    <value>10211</value>

                    //Removed 07/Apr/2022
                    <condition attribute='createdon' operator='on-or-after' value='" +_fromDate + @"' />
                    <condition attribute='createdon' operator='on-or-before' value='" + _toDate + @"' />
                    */

                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='annotation'>
                        <attribute name='objectid' />
                        <attribute name='filename' />
                        <attribute name='createdon' />
                        <attribute name='filesize' />
                        <attribute name='objecttypecode' />
                        <attribute name='isdocument' />
                        <order attribute='createdon' descending='false' />
                        <filter type='and'>
                            <condition attribute='filesize' operator='gt' value='0' />
                            <condition attribute='objectid' operator='not-null' />
                            <condition attribute='objecttypecode' operator='not-in'>
                                <value>10211</value>
                            </condition> 
                        </filter>
                        </entity>
                        </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    i = 0;
                    j = EntityList.Entities.Count;
                    foreach (Entity entity in EntityList.Entities)
                    {
                        Console.WriteLine("Date:" + entity.GetAttributeValue<DateTime>("createdon").ToString());
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName.IndexOf("adx") < 0)
                        {
                            if (EntityLogReference.LogicalName == "msdyn_survey" ||  EntityLogReference.LogicalName == "hil_tender" || EntityLogReference.LogicalName == "contact" || EntityLogReference.LogicalName == "msdyn_workordersubstatus" || EntityLogReference.LogicalName == "hil_sawactivityapproval" || EntityLogReference.LogicalName == "hil_technician" || EntityLogReference.LogicalName == "product" || EntityLogReference.LogicalName == "msdyn_incidenttype" || EntityLogReference.LogicalName == "hil_unitwarranty" || EntityLogReference.LogicalName == "msdyn_approval" || EntityLogReference.LogicalName == "hil_feedback" || EntityLogReference.LogicalName == "hil_inventoryrequest" || EntityLogReference.LogicalName == "msdyn_workorderservice" || EntityLogReference.LogicalName == "msdyn_customerasset" || EntityLogReference.LogicalName == "msdyn_workorder" || EntityLogReference.LogicalName == "msdyn_workorderproduct")
                            {
                                filename = entity.GetAttributeValue<string>("filename");
                                notetext = entity.GetAttributeValue<string>("notetext");
                                createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                                filesize = entity.GetAttributeValue<int>("filesize");
                                objectid = entity.GetAttributeValue<EntityReference>("objectid");
                                isDocument = entity.GetAttributeValue<bool>("isdocument");
                                _blobURI = string.Empty;
                                string _refId = string.Empty;
                                if (entity.Attributes.Contains("ae.msdyn_name"))
                                {
                                    _refId = entity.GetAttributeValue<AliasedValue>("ae.msdyn_name").Value.ToString();
                                }
                                if (objectid != null && filesize > 0)
                                {
                                    entTemp = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody"));
                                    if (entTemp.Contains("documentbody") || entTemp.Attributes.Contains("documentbody"))
                                    {
                                        documentbody = entTemp.GetAttributeValue<string>("documentbody");
                                        entity["subject"] = filename;
                                        entity["documentbody"] = null;
                                        entity["filename"] = null;
                                        entity["filesize"] = 0;
                                        filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                        try
                                        {
                                            //msdyn_workordersubstatus
                                            if (EntityLogReference.LogicalName == "msdyn_workorderservice" ||  EntityLogReference.LogicalName == "msdyn_workorderservice" || EntityLogReference.LogicalName == "msdyn_workorder" || EntityLogReference.LogicalName == "msdyn_workorderproduct")
                                            {
                                                _blobURI = Upload(filename, documentbody, BlobContainers.WORKORDER);
                                            }
                                            if (EntityLogReference.LogicalName == "msdyn_customerasset")
                                            {
                                                _blobURI = Upload(filename, documentbody, BlobContainers.CUSTOMERASSET);
                                            }
                                            else
                                            {
                                                _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                            }
                                            entity["notetext"] = _blobURI;
                                            _service.Update(entity);
                                            Console.WriteLine(EntityLogReference.LogicalName + "#" + _refId + " : " + i.ToString() + "/" + j.ToString());
                                        }
                                        catch
                                        {
                                            Console.WriteLine("Error !! Attachment is not proper base64." + "/" + j.ToString());
                                        }
                                    }
                                } 
                            }
                            else {
                                Console.WriteLine(EntityLogReference.LogicalName + i.ToString() + "/" + j.ToString());
                            }
                        }
                        else {
                            Console.WriteLine(EntityLogReference.LogicalName + i.ToString() + "/" + j.ToString());
                        }
                        i += 1;
                        Console.WriteLine(EntityLogReference.LogicalName + i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ":FileName="+ filename);
            }
        }
        static void CustomerAsset()
        {
            //await Task.Run(() =>
            //{
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='annotation'>
                        <attribute name='objectid' />
                        <attribute name='filename' />
                        <attribute name='createdon' />
                        <attribute name='filesize' />
                        <attribute name='objecttypecode' />
                        <attribute name='isdocument' />
                        <order attribute='createdon' descending='false' />
                        <filter type='and'>
                            <condition attribute='filesize' operator='gt' value='0' />
                        </filter>
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='objectid' link-type='inner' alias='ad' />
                        </entity>
                        </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_customerasset")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.CUSTOMERASSET);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Customer Asset Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //});
        }
        static void DownloadCustomerSignature() {
            try
            {
                string query = String.Format(@"<fetch mapping='logical'>
                <entity name='{0}'>
                <attribute name='msdyn_name' />
                <attribute name='entityimage' />
                <attribute name='createdon' />
                <order attribute='createdon' descending='false' />
                <filter type='and'>
                    <condition attribute='entityimage' operator='not-null' />
                </filter>
                </entity>
                </fetch>", "msdyn_workorder");

                EntityCollection images = _service.RetrieveMultiple(new FetchExpression(query));
                int i=0, j = 0;
                j = images.Entities.Count;

                foreach (Entity record in images.Entities)
                {
                    string photoName = "CustomerSignature_" + record["msdyn_name"] + ".jpg" as string;
                    string bitString = Convert.ToBase64String(record["entityimage"] as byte[]);
                    byte[] NoteByte = record["entityimage"] as byte[];
                    try
                    {
                        string _blobURI = Upload(photoName, bitString, BlobContainers.WORKORDER);

                        Annotation An = new Annotation();
                        An.DocumentBody = null;
                        An.FileName = null;
                        An["filesize"] = 0;
                        An.Subject = photoName;
                        An.NoteText = _blobURI;
                        An.ObjectId = new EntityReference(msdyn_workorder.EntityLogicalName, record.Id);
                        An.ObjectTypeCode = msdyn_workorder.EntityLogicalName;
                        try
                        {
                            _service.Create(An);
                            Entity entImage = new Entity(msdyn_workorder.EntityLogicalName);
                            entImage.Id = record.Id;
                            entImage["entityimage"] = null;
                            _service.Update(entImage);
                        }
                        catch {}
                        i++;
                        Console.WriteLine(photoName + " :" + i.ToString() + "/" + j.ToString());
                    }
                    catch
                    {
                        Console.WriteLine("Error !! Attachment is not proper base64.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle exception
            }

        }
        static CrmServiceClient GetCRMConnection()
        {
            CrmServiceClient service = null;
            try
            {
                string authType = "OAuth";
                string userName = "kuldeep.khare@smylsolutions.com";
                string password = "Patanahi@1234";
                string url = "https://org72dd2e85.crm11.dynamics.com/";
                string appId = "51f81489-12ee-4a9e-aaae-a2591f45987d";
                string reDirectURI = "app://58145B91-0C36-4500-8554-080854F2AC97";
                string loginPrompt = "Auto";
                string ConnectionString = string.Format("AuthType = {0};Username = {1};Password = {2}; Url = {3}; AppId={4}; RedirectUri={5};LoginPrompt={6}",
                                                        authType, userName, password, url, appId, reDirectURI, loginPrompt);

                service = new CrmServiceClient(ConnectionString);
                loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
            }
            catch { }
            return service;
        }
        static void DownloadCustomerAssetImage()
        {
            try
            {
                string query = String.Format(@"<fetch mapping='logical'>
                    <entity name='{0}'>
                    <attribute name='msdyn_name' />
                    <attribute name='entityimage' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='entityimage' operator='not-null' />
                    </filter>
                    </entity>
                </fetch>", "msdyn_customerasset");

                while (true)
                {
                    EntityCollection images = _service.RetrieveMultiple(new FetchExpression(query));
                    int i = 0, j = 0;
                    j = images.Entities.Count;
                    if (j == 0) { break; }
                    foreach (Entity record in images.Entities)
                    {
                        DateTime _createdOn = record.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                        string _serialNumber = record.GetAttributeValue<string>("msdyn_name");

                        string photoName = "AssetImage_" + record.Id.ToString() + ".jpg" as string;
                        string bitString = Convert.ToBase64String(record["entityimage"] as byte[]);
                        byte[] NoteByte = record["entityimage"] as byte[];
                        try
                        {
                            string _blobURI = Upload(photoName, bitString, BlobContainers.CUSTOMERASSET);

                            Annotation An = new Annotation();
                            An.DocumentBody = null;
                            An.FileName = null;
                            An["filesize"] = 0;
                            An.Subject = "AssetImage";
                            An.NoteText = _blobURI;
                            An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, record.Id);
                            An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                            try
                            {
                                _service.Create(An);
                                Entity entImage = new Entity(msdyn_customerasset.EntityLogicalName);
                                entImage.Id = record.Id;
                                entImage["entityimage"] = null;
                                _service.Update(entImage);
                            }
                            catch { }
                            i++;
                            Console.WriteLine(_createdOn.ToString() + "/" + _serialNumber + " :" + i.ToString() + "/" + j.ToString());
                        }
                        catch
                        {
                            Console.WriteLine("Error !! Attachment is not proper base64.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        static void Customer()
        {
            //await Task.Run(() =>
            //{
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='annotation'>
                        <attribute name='objectid' />
                        <attribute name='filename' />
                        <attribute name='createdon' />
                        <attribute name='filesize' />
                        <attribute name='objecttypecode' />
                        <attribute name='isdocument' />
                        <order attribute='createdon' descending='false' />
                        <link-entity name='contact' from='contactid' to='objectid' link-type='inner' alias='am' />
                        </entity>
                        </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "contact")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.CUSTOMERASSET);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Customer Asset Attachment Migrated: " + i.ToString());
                            }
                            else {
                                _service.Delete("annotation", entity.Id);
                                i += 1;
                                Console.WriteLine("Attachment Deleted: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //});
        }
        static void WithOutRegarding()
        {
            //await Task.Run(() =>
            //{
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='annotation'>
                        <attribute name='objectid' />
                        <attribute name='filename' />
                        <attribute name='createdon' />
                        <attribute name='filesize' />
                        <attribute name='objecttypecode' />
                        <attribute name='isdocument' />
                        <order attribute='createdon' descending='false' />
                        <filter type='and'>
                            <condition attribute='filesize' operator='gt' value='0' />
                            <condition attribute='objectid' operator='null' />
                            <condition attribute='createdon' operator='on-or-before' value='2019-03-31' />
                            <filter type='or'>
                                <condition attribute='notetext' operator='like' value='%serial%' />
                                <condition attribute='notetext' operator='like' value='%invoice%' />
                            </filter>
                        </filter>
                        </entity>
                        </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        string filename = entity.GetAttributeValue<string>("filename");
                        DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                        int filesize = entity.GetAttributeValue<int>("filesize");
                        bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                        string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                        string _blobURI = string.Empty;
                        if (isDocument && filesize > 0 && documentbody != null)
                        {
                            entity["subject"] = filename;
                            entity["documentbody"] = null;
                            entity["filename"] = null;
                            entity["filesize"] = 0;
                            filename =  "WithoutRef_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                            try
                            {
                                _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                entity["notetext"] = _blobURI;
                                _service.Update(entity);
                            }
                            catch
                            {
                                Console.WriteLine("Error !! Attachment is not proper base64.");
                            }
                            i += 1;
                            Console.WriteLine("Without Ref Attachment Migrated: " + i.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //});
        }
        static void WorkOrder()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='objectid' link-type='inner' alias='ae' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_workorder")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.WORKORDER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Work Order Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void WorkOrderProduct()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='msdyn_workorderproduct' from='msdyn_workorderproductid' to='objectid' link-type='inner' alias='ai' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_workorderproduct")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody,BlobContainers.WORKORDER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Work Order Product Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void WorkOrderService()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='msdyn_workorderservice' from='msdyn_workorderserviceid' to='objectid' link-type='inner' alias='ai' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_workorderservice")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.WORKORDER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Work Order Service Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void InventoryRequest()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                string notetext=string.Empty;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='notetext' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <link-entity name='hil_inventoryrequest' from='hil_inventoryrequestid' to='objectid' link-type='inner' alias='bo' />
                    </entity>
                    </fetch>";

                    /*<filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    */

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "hil_inventoryrequest")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            notetext = entity.GetAttributeValue<string>("notetext");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Inventory Request Attachment Migrated: " + i.ToString());
                            }
                            else
                            {
                                if (notetext == null || notetext.IndexOf("d365storagesa.blob") < 0)
                                {
                                    _service.Delete("annotation", entity.Id);
                                    i += 1;
                                    Console.WriteLine("Attachment Deleted: " + i.ToString());
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void ConsumerFeedback()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='hil_feedback' from='hil_feedbackid' to='objectid' link-type='inner' alias='ca' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "hil_feedback")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Consumer Feedback Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void PRApprovals()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='msdyn_approval' from='activityid' to='objectid' link-type='inner' alias='ce' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_approval")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("PR Approvals Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void UnitWarranty()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='hil_unitwarranty' from='hil_unitwarrantyid' to='objectid' link-type='inner' alias='cg' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "hil_unitwarranty")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Unit Warranty Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void IncidentTypeCause()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='objectid' link-type='inner' alias='cj' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_incidenttype")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Incident Type (Cause) Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static void Product()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='annotation'>
                    <attribute name='objectid' />
                    <attribute name='filename' />
                    <attribute name='createdon' />
                    <attribute name='filesize' />
                    <attribute name='objecttypecode' />
                    <attribute name='isdocument' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='filesize' operator='gt' value='0' />
                    </filter>
                    <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='objectid' link-type='inner' alias='cj' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    foreach (Entity entity in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)entity.Attributes["objectid"]);
                        if (EntityLogReference.LogicalName == "msdyn_incidenttype")
                        {
                            string filename = entity.GetAttributeValue<string>("filename");
                            DateTime createdon = entity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            int filesize = entity.GetAttributeValue<int>("filesize");
                            EntityReference objectid = entity.GetAttributeValue<EntityReference>("objectid");
                            bool isDocument = entity.GetAttributeValue<bool>("isdocument");
                            string documentbody = _service.Retrieve("annotation", entity.Id, new ColumnSet("documentbody")).GetAttributeValue<string>("documentbody");
                            string _blobURI = string.Empty;
                            if (isDocument && objectid != null && filesize > 0 && documentbody != null)
                            {
                                entity["subject"] = filename;
                                entity["documentbody"] = null;
                                entity["filename"] = null;
                                entity["filesize"] = 0;
                                filename = EntityLogReference.LogicalName + "_" + entity.Id + "_" + createdon.ToString() + "_" + filename;
                                try
                                {
                                    _blobURI = Upload(filename, documentbody, BlobContainers.OTHER);
                                    entity["notetext"] = _blobURI;
                                    _service.Update(entity);
                                }
                                catch
                                {
                                    Console.WriteLine("Error !! Attachment is not proper base64.");
                                }
                                i += 1;
                                Console.WriteLine("Incident Type (Cause) Attachment Migrated: " + i.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        static string ToURLSlug(string s)
        {
            return Regex.Replace(s, @"[^a-z0-9]+", "-", RegexOptions.IgnoreCase)
                .Trim(new char[] { '-' })
                .ToLower();
        }
        static string Upload(string fileName, string noteBody,string containerName)
        {
            string _blobURI = string.Empty;
            try
            {
                if (noteBody.IndexOf("base64,") > 0)
                {
                    noteBody = noteBody.Substring(noteBody.IndexOf("base64,") + 7);
                }
                byte[] fileContent = Convert.FromBase64String(noteBody);
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
                throw new Exception("Havells_Plugin.Annotations.PreCreate.Upload: " + ex.Message);
            }
            return _blobURI;
        }

        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }

        static bool DeleteBlobFile() {
            bool _retValue = false;
            try
            {
                var _containerName = "homeadvisory";
                string _storageConnection = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net"; ;
                CloudStorageAccount cloudStorageAccount = CloudStorageAccount.Parse(_storageConnection);
                CloudBlobClient _blobClient = cloudStorageAccount.CreateCloudBlobClient();
                CloudBlobContainer _cloudBlobContainer = _blobClient.GetContainerReference(_containerName);
                CloudBlockBlob _blockBlob = _cloudBlobContainer.GetBlockBlobReference("23B5196D-0E7B-EB11-A812-0022486EAB14_Home Advisory_202103020420577483.xml");
                _blockBlob.Delete();
                _retValue = true;
            }
            catch {
            }
            return _retValue;
        }
        #endregion
    }
    public class AuthenticateConsumer
    {
        public string JWTToken { get; set; }
        public string MobileNumber { get; set; }
        public string SourceOrigin { get; set; }
        public string SessionId { get; set; }

        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }

    }
}
