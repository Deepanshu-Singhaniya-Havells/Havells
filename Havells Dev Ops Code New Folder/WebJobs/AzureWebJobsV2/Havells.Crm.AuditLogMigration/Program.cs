using Microsoft.Crm.Sdk.Messages;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Havells.Crm.AuditLogMigration
{
    class Program
    {
        public static Dictionary<string, int?> AuditLogEnabledEntityList = new Dictionary<string, int?>();
        public static int[] notLoggedentity = {
            9007, 9983, 3005, 4720, 10546, 10547, 10548, 10832, 11299, 11311, 11331, 10760, 10761, 10762, 10766, 10825, 10672, 10578, //10144//, 10196
        };
        static void Main(string[] args)
        {
            //https://havells.crm8.dynamics.com/api/data/v9.0/EntityDefinitions

            var fromDate = ConfigurationManager.AppSettings["fromDate"].ToString();
            var toDate = ConfigurationManager.AppSettings["toDate"].ToString();
            var objecttypeCode = ConfigurationManager.AppSettings["ObjectTypeCode"].ToString();

            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = createConnection(finalString);

            //ListAllAuditEnableEntity(service);

            ////Delete Transactional Data
            //DeleteTransactionalData(service, "msdyn_workorderservice", fromDate, toDate);

            //GetAuditLogHeader(service, int.Parse(objecttypeCode), fromDate, toDate);

            GetAuditLogByEntityDailyLogs(service, int.Parse(objecttypeCode), fromDate, toDate);

            //GetAllAuditEnableEntity(service, fromDate, toDate);

            //updateFlag(entityName, fromDate, toDate, service);
            //D365AuditLogMigration_DeltaWithFlag("msdyn_customerasset", "2021-07-01", "2021-07-31", service);

            //GetAuditEnableEntity(service);
            //foreach (KeyValuePair<string, int?> kvp in AuditLogEnabledEntityList)
            //{
            //    Console.WriteLine("Key = {0}, Value = {1}", kvp.Key, kvp.Value);
            //    //D365AuditLogMigration_Delta(kvp.Key, service);
            //}
        }

        public static void GetAllAuditEnableEntity(IOrganizationService service, string fromDate, string toDate)
        {
            try
            {
                RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
                RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
                metaDataRequest.EntityFilters = EntityFilters.Entity;

                XmlDictionaryReaderQuotas myReaderQuotas = new XmlDictionaryReaderQuotas();
                myReaderQuotas.MaxNameTableCharCount = 2147483647;

                // Execute the request.

                metaDataResponse = (RetrieveAllEntitiesResponse)service.Execute(metaDataRequest);

                var entities = metaDataResponse.EntityMetadata;
                foreach (var abc in entities)
                {
                    try
                    {
                        if (abc.ObjectTypeCode == 10460 || abc.ObjectTypeCode == 10695 || abc.ObjectTypeCode == 11222
                            || abc.ObjectTypeCode == 10972 || abc.ObjectTypeCode == 11223 || abc.ObjectTypeCode == 9007 
                            || abc.ObjectTypeCode == 9983 || abc.ObjectTypeCode == 3005 || abc.ObjectTypeCode == 11401 || abc.ObjectTypeCode == 4720
                            || abc.ObjectTypeCode == 10672 || abc.ObjectTypeCode == 11458 || abc.ObjectTypeCode == 11454 || abc.ObjectTypeCode == 10578

                            )
                            continue;
                        //int abccc = Array.Find(notLoggedentity, element => element == abc.ObjectTypeCode);
                        if (abc.IsAuditEnabled.Value == true)//&& abccc == 0)
                        {
                            Console.WriteLine("++++++++++++++++++++++++++++ Entity Name " + abc.LogicalName + " ++++++++++++++++++++++++++++");
                            int? objecttypeCode = abc.ObjectTypeCode;
                            GetAuditLogHeader(service, objecttypeCode, fromDate, toDate);
                            Console.WriteLine("++++++++++++++++++++++++++++ Done ++++++++++++++++++++++++++++");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error in Loop : " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
        }
        public static void ListAllAuditEnableEntity(IOrganizationService service)
        {
            RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
            RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
            metaDataRequest.EntityFilters = EntityFilters.Entity;

            XmlDictionaryReaderQuotas myReaderQuotas = new XmlDictionaryReaderQuotas();
            myReaderQuotas.MaxNameTableCharCount = 2147483647;

            // Execute the request.

            metaDataResponse = (RetrieveAllEntitiesResponse)service.Execute(metaDataRequest);

            var entities = metaDataResponse.EntityMetadata;
            foreach (var abc in entities)
            {
                //if (abc.ObjectTypeCode == 10460 || abc.ObjectTypeCode == 10695)
                //    continue;
                //int abccc = Array.Find(notLoggedentity, element => element == abc.ObjectTypeCode);
                if (abc.IsAuditEnabled.Value == true)//&& abccc == 0)
                {
                    Console.WriteLine(abc.LogicalName + "|" + abc.ObjectTypeCode + "|" + getDate(service, abc.ObjectTypeCode));
                    //int? objecttypeCode = abc.ObjectTypeCode;
                    //GetAuditLogHeader(service, objecttypeCode, fromDate, toDate);

                }
            }
        }
        public static string getDate(IOrganizationService service, int? objectTypeCode)
        {
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            List<EntityAttribute> lstEntityAttribute = new List<EntityAttribute>();
            //Console.WriteLine("Record Migrating of Entity Code: " + objectTypeCode + " from: " + fromDate + " till: " + toDate);
            List<Guid> guids = new List<Guid>();
            string entityLogicalName = string.Empty;
            int pageNo = 1;
            int totalRecCount = 0;
            int totalAuditRecCount = 0;
            int totalRecDone = 0;

            string accName = "d365storagesa";
            string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            //bool _retValue = false;
            string _recID = string.Empty;

            // Implement the accout, set true for https for SSL.  
            StorageCredentials creds = new StorageCredentials(accName, accKey);
            CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
            // Create the blob client.  
            CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
            // Retrieve a reference to a container.   
            CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
            // Create the container if it doesn't already exist.  
            container.CreateIfNotExistsAsync();


        repetRetrive:
            guids = new List<Guid>();

            string fetchXml = $@"<fetch top=""1"" version='1.0' mapping='logical' >
                    <entity name='audit' >   
                    <attribute name='operation' />
                    <attribute name='objecttypecodename' />
                    <attribute name='objectid' />
                    <attribute name='operationname' />
                    <attribute name='transactionid' />
                    <attribute name='useradditionalinfo' />
                    <attribute name='regardingobjectidname' />
                    <attribute name='auditid' />
                    <attribute name='createdon' />
                    <attribute name='userid' />
                    <attribute name='callinguseridname' />
                    <attribute name='regardingobjectid' />
                    <attribute name='objecttypecode' />
                    <attribute name='action' />
                    <attribute name='actionname' />
                    <attribute name='objectidname' />
                    <attribute name='callinguserid' />
                    <order attribute='createdon' descending='false' />
                        <filter type='and'>
                        <condition attribute='objecttypecode' operator='eq' value='{objectTypeCode}' />
                    </filter>
                    </entity>
            </fetch>";

            EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));
            totalRecCount = entityCollection.Entities.Count + totalRecCount;
            if (totalRecCount > 0)
                return entityCollection[0].GetAttributeValue<DateTime>("createdon").ToString("dd/MM/yyyy");
            else
                return null;

        }
        public static void DeleteTransactionalData(IOrganizationService service, string objectTypeCode, string fromDate, string toDate)
        {
            Console.WriteLine("Record Migrating of Entity Name: " + objectTypeCode + " from: " + fromDate + " till: " + toDate);
            List<Guid> guids = new List<Guid>();
            string entityLogicalName = string.Empty;
            int pageNo = 1;
            int totalRecCount = 0;
            int totalAuditRecCount = 0;
            int totalRecDone = 0;

        repetRetrive:
            guids = new List<Guid>();

            string fetchXml = $@"<fetch page='" + pageNo + $@"' version='1.0' mapping='logical' >
                        <entity name='{objectTypeCode}' >   
                        <order attribute='createdon' descending='false' />
                            <filter type='and'>
                            <condition attribute='createdon' operator='on-or-before' value='2021-04-30' />
                        </filter>
                        </entity>
                </fetch>";
            EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));
            totalRecCount = entityCollection.Entities.Count + totalRecCount;

            Console.WriteLine("From " + fromDate + " To " + toDate + " Count " + entityCollection.Entities.Count);

            foreach (Entity entity in entityCollection.Entities)
            {
                Guid objId = entity.Id;
                Guid abccc = Array.Find(guids.ToArray(), element => element == objId);
                if (abccc == Guid.Empty && objId != Guid.Empty)
                {
                    guids.Add(objId);
                }
            }

            totalAuditRecCount = totalAuditRecCount + guids.Count;
            Console.WriteLine("Entity Name: " + objectTypeCode + " JOb Count is " + totalAuditRecCount);

            //Console.WriteLine("**************************************");
            Parallel.ForEach(guids, entityId =>
            {
                totalRecDone++;
                int i = totalRecDone;
                Console.WriteLine(" #################################### Record Started " + i + "/" + totalAuditRecCount + " #################################### ");
                service.Delete(objectTypeCode, entityId);
                Console.WriteLine(" #################################### Record Completed " + i + "/" + totalAuditRecCount + " #################################### ");
                //Console.WriteLine,
            });

            if (entityCollection.Entities.Count == 500)
            {
                pageNo++;
                goto repetRetrive;
            }
            Console.WriteLine("**************************************");
            Console.WriteLine("Record Migrated of Entity Name: " + entityLogicalName + " from " + fromDate + " till " + toDate);
        }
        public static void GetAuditLogByEntityDailyLogs(IOrganizationService service, int? objectTypeCode, string fromDate, string toDate)
        {
            //string fromDate = DateTime.Now.ToString("yyyy-MM-dd");
            //string toDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
            //string toDate;
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            List<EntityAttribute> lstEntityAttribute = new List<EntityAttribute>();
            Console.WriteLine("Record Migrating of Entity Code: " + objectTypeCode.ToString() + " from: " + fromDate + " till: " + toDate);
            List<Guid> guidList = new List<Guid>();
            string entityLogicalName = string.Empty;
            int num1 = 1;
            int num2 = 0;
            int totalAuditRecCount = 0;
            int totalRecDone = 0;
            string accountName = "d365storagesa";
            string keyValue = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            string empty = string.Empty;
            CloudBlobContainer container = new CloudStorageAccount(new StorageCredentials(accountName, keyValue), true).CreateCloudBlobClient().GetContainerReference("d365auditlogarchive");
            container.CreateIfNotExistsAsync();
            while (true)
            {
                List<Guid> source = new List<Guid>();
                string query = "<fetch page='" + num1.ToString() + string.Format("' version='1.0' mapping='logical' >" +
                    "<entity name='audit' >" +
                    "<attribute name='operation' />" +
                    "<attribute name='objecttypecodename' />" +
                    "<attribute name='objectid' />" +
                    "<attribute name='operationname' />" +
                    "<attribute name='transactionid' />" +
                    "<attribute name='useradditionalinfo' />" +
                    "<attribute name='regardingobjectidname' />" +
                    "<attribute name='auditid' />" +
                    "<attribute name='createdon' />" +
                    "<attribute name='userid' />" +
                    "<attribute name='callinguseridname' />" +
                    "<attribute name='regardingobjectid' />" +
                    "<attribute name='objecttypecode' />" +
                    "<attribute name='action' />" +
                    "<attribute name='actionname' />" +
                    "<attribute name='objectidname' />" +
                    "<attribute name='callinguserid' />" +
                    "<order attribute='createdon' descending='false' />" +
                    "<filter type='and'>" +
                    "<condition attribute='objecttypecode' operator='eq' value='{0}' />" +
                    "<condition attribute='createdon' operator='gt' value='{1}' />" +
                    "<condition attribute='createdon' operator='lt' value='{2}' />" +
                    "<condition attribute='action' operator='in'>" +
                    "<value>1</value>" +
                    "<value>2</value><value>4</value>" +
                    "<value>5</value>" +
                    "<value>13</value>" +
                    "</condition>" +
                    "</filter>" +
                    "</entity>" +
                    "</fetch>", (object)objectTypeCode, (object)fromDate, (object)toDate);
                EntityCollection entityCollection = service.RetrieveMultiple((QueryBase)new FetchExpression(query));
                num2 = entityCollection.Entities.Count + num2;
                Console.WriteLine("From " + fromDate + " To " + toDate + " Count " + entityCollection.Entities.Count.ToString());
                foreach (Entity entity in (Collection<Entity>)entityCollection.Entities)
                {
                    Guid objId = entity.GetAttributeValue<EntityReference>("objectid").Id;
                    if (Array.Find<Guid>(source.ToArray(), (Predicate<Guid>)(element => element == objId)) == Guid.Empty && objId != Guid.Empty)
                        source.Add(objId);
                    if (entityLogicalName == string.Empty)
                    {
                        entityLogicalName = entity.GetAttributeValue<EntityReference>("objectid").LogicalName;
                        Program.GetPrimaryIdFieldName(entityLogicalName, out _primaryKey, out _primaryfield, service);
                        lstEntityAttribute = Program.GetEntityAttributes(entityLogicalName, service);
                    }
                }
                totalAuditRecCount += source.Count;
                Console.WriteLine("Entity Name: " + entityLogicalName + " JOb Count is " + totalAuditRecCount.ToString());
                Parallel.ForEach<Guid>((IEnumerable<Guid>)source, (Action<Guid>)(entityId =>
                {
                    ++totalRecDone;
                    int num3 = totalRecDone;
                    Console.WriteLine(" #################################### Record Started " + num3.ToString() + "/" + totalAuditRecCount.ToString() + " #################################### ");
                    Console.WriteLine((object)entityId);
                    Program.migrateAuditLogs(service, entityLogicalName, entityId, _primaryfield, _primaryKey, lstEntityAttribute, container);
                    Console.WriteLine(" #################################### Record Completed " + num3.ToString() + "/" + totalAuditRecCount.ToString() + " #################################### ");
                }));
                if (entityCollection.Entities.Count > 0)
                    ++num1;
                else
                    break;
            }
            Console.WriteLine("**************************************");
            Console.WriteLine("Record Migrated of Entity Name: " + entityLogicalName + " from " + fromDate + " till " + fromDate);
        }
        public static void GetAuditLogHeader(IOrganizationService service,int? objectTypeCode,string fromDate,string toDate)
        {
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            List<EntityAttribute> lstEntityAttribute = new List<EntityAttribute>();
            Console.WriteLine("Record Migrating of Entity Code: " + objectTypeCode.ToString() + " from: " + fromDate + " till: " + toDate);
            List<Guid> guidList = new List<Guid>();
            string entityLogicalName = string.Empty;
            int num1 = 1;
            int num2 = 0;
            int totalAuditRecCount = 0;
            int totalRecDone = 0;
            string accountName = "d365storagesa";
            string keyValue = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            string empty = string.Empty;
            CloudBlobContainer container = new CloudStorageAccount(new StorageCredentials(accountName, keyValue), true).CreateCloudBlobClient().GetContainerReference("d365auditlogarchive");
            container.CreateIfNotExistsAsync();
            while (true)
            {
                List<Guid> source = new List<Guid>();
                string query = "<fetch page='" + num1.ToString() + string.Format("' version='1.0' mapping='logical' >" +
                    "<entity name='audit' >" +
                    "<attribute name='operation' />" +
                    "<attribute name='objecttypecodename' />" +
                    "<attribute name='objectid' />" +
                    "<attribute name='operationname' />" +
                    "<attribute name='transactionid' />" +
                    "<attribute name='useradditionalinfo' />" +
                    "<attribute name='regardingobjectidname' />" +
                    "<attribute name='auditid' />" +
                    "<attribute name='createdon' />" +
                    "<attribute name='userid' />" +
                    "<attribute name='callinguseridname' />" +
                    "<attribute name='regardingobjectid' />" +
                    "<attribute name='objecttypecode' />" +
                    "<attribute name='action' />" +
                    "<attribute name='actionname' />" +
                    "<attribute name='objectidname' />" +
                    "<attribute name='callinguserid' />" +
                    "<order attribute='createdon' descending='false' />" +
                    "<filter type='and'>" +
                    "<condition attribute='objecttypecode' operator='eq' value='{0}' />" +
                    "<condition attribute='createdon' operator='gt' value='{1}' />" +
                    "<condition attribute='createdon' operator='lt' value='{2}' />" +
                    "<condition attribute='action' operator='in'>" +
                    "<value>1</value>" +
                    "<value>2</value><value>4</value>" +
                    "<value>5</value>" +
                    "<value>13</value>" +
                    "</condition>" +
                    "</filter>" +
                    "</entity>" +
                    "</fetch>", (object)objectTypeCode, (object)fromDate, (object)toDate);
                EntityCollection entityCollection = service.RetrieveMultiple((QueryBase)new FetchExpression(query));
                num2 = entityCollection.Entities.Count + num2;
                Console.WriteLine("From " + fromDate + " To " + toDate + " Count " + entityCollection.Entities.Count.ToString());
                foreach (Entity entity in (Collection<Entity>)entityCollection.Entities)
                {
                    Guid objId = entity.GetAttributeValue<EntityReference>("objectid").Id;
                    if (Array.Find<Guid>(source.ToArray(), (Predicate<Guid>)(element => element == objId)) == Guid.Empty && objId != Guid.Empty)
                        source.Add(objId);
                    if (entityLogicalName == string.Empty)
                    {
                        entityLogicalName = entity.GetAttributeValue<EntityReference>("objectid").LogicalName;
                        Program.GetPrimaryIdFieldName(entityLogicalName, out _primaryKey, out _primaryfield, service);
                        lstEntityAttribute = Program.GetEntityAttributes(entityLogicalName, service);
                    }
                }
                totalAuditRecCount += source.Count;
                Console.WriteLine("Entity Name: " + entityLogicalName + " JOb Count is " + totalAuditRecCount.ToString());
                Parallel.ForEach<Guid>((IEnumerable<Guid>)source, (Action<Guid>)(entityId =>
                {
                    ++totalRecDone;
                    int num3 = totalRecDone;
                    Console.WriteLine(" #################################### Record Started " + num3.ToString() + "/" + totalAuditRecCount.ToString() + " #################################### ");
                    Console.WriteLine((object)entityId);
                    Program.migrateAuditLogs(service, entityLogicalName, entityId, _primaryfield, _primaryKey, lstEntityAttribute, container);
                    Console.WriteLine(" #################################### Record Completed " + num3.ToString() + "/" + totalAuditRecCount.ToString() + " #################################### ");
                }));
                if (entityCollection.Entities.Count > 0)
                    ++num1;
                else
                    break;
            }
            Console.WriteLine("**************************************");
            Console.WriteLine("Record Migrated of Entity Name: " + entityLogicalName + " from " + fromDate + " till " + toDate);
        }

        private static void migrateAuditLogs(
      IOrganizationService _service,
      string entityLogicalName,
      Guid entityId,
      string _primaryfield,
      string _primaryKey,
      List<EntityAttribute> lstEntityAttribute,
      CloudBlobContainer container)
        {
            string[] strArray1 = new string[3]
            {
        _primaryKey,
        _primaryfield,
        "createdon"
            };
            string empty1 = string.Empty;
            try
            {
                empty1 = entityId.ToString();
                StringBuilder stringBuilder = new StringBuilder();
                if (!Program.CheckIfAzureBlobExist(entityId.ToString(), container))
                    stringBuilder.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                RetrieveRecordChangeHistoryRequest request1 = new RetrieveRecordChangeHistoryRequest()
                {
                    Target = new EntityReference(entityLogicalName, entityId)
                };
                request1.Target.LogicalName = entityLogicalName;
                RetrieveRecordChangeHistoryResponse changeHistoryResponse;
                while (true)
                {
                    try
                    {
                        changeHistoryResponse = (RetrieveRecordChangeHistoryResponse)_service.Execute((OrganizationRequest)request1);
                        break;
                    }
                    catch
                    {
                    }
                }
                AuditDetailCollection detailCollection = changeHistoryResponse.AuditDetailCollection;
                bool flag = false;
                foreach (AuditDetail auditDetail in (Collection<AuditDetail>)detailCollection.AuditDetails)
                {
                    Entity auditRecord = auditDetail.AuditRecord;
                    string empty2 = string.Empty;
                    string empty3 = string.Empty;
                    string formattedValue = auditRecord.FormattedValues["action"];
                    if (auditRecord.Attributes.Contains("objectid") && (formattedValue == "Create" || formattedValue == "Update" || formattedValue == "Delete" || formattedValue == "Activate" || formattedValue == "Deactivate" || formattedValue == "Assign"))
                    {
                        string str1 = !auditRecord.Attributes.Contains("auditid") ? (string)null : auditRecord.Attributes["auditid"].ToString();
                        int num1;
                        string str2;
                        if (auditRecord.Attributes.Contains("createdon"))
                        {
                            DateTime dateTime = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330.0);
                            string[] strArray2 = new string[8];
                            strArray2[0] = str1;
                            strArray2[1] = ",";
                            num1 = dateTime.Year;
                            strArray2[2] = num1.ToString();
                            num1 = dateTime.Month;
                            strArray2[3] = num1.ToString().PadLeft(2, '0');
                            num1 = dateTime.Day;
                            strArray2[4] = num1.ToString().PadLeft(2, '0');
                            num1 = dateTime.Hour;
                            strArray2[5] = num1.ToString().PadLeft(2, '0');
                            num1 = dateTime.Minute;
                            strArray2[6] = num1.ToString().PadLeft(2, '0');
                            num1 = dateTime.Second;
                            strArray2[7] = num1.ToString().PadLeft(2, '0');
                            str2 = string.Concat(strArray2);
                        }
                        else
                            str2 = str1 + "," + string.Empty;
                        string str3;
                        if (auditRecord.Attributes.Contains("action"))
                        {
                            string str4 = str2;
                            num1 = ((OptionSetValue)auditRecord.Attributes["action"]).Value;
                            string str5 = num1.ToString();
                            str3 = str4 + "," + str5 + "," + auditRecord.FormattedValues["action"];
                        }
                        else
                            str3 = str2 + "," + string.Empty + "," + string.Empty;
                        string str6 = !auditRecord.Attributes.Contains("objectid") ? str3 + "," + string.Empty : str3 + "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                        string str7;
                        if (auditRecord.Attributes.Contains("objecttypecode"))
                            str7 = str6 + "," + auditRecord.Attributes["objecttypecode"].ToString() + "," + auditRecord.Attributes["objecttypecode"].ToString();
                        else
                            str7 = str6 + "," + string.Empty + "," + string.Empty;
                        string str8;
                        if (auditRecord.Attributes.Contains("operation"))
                        {
                            string str9 = str7;
                            num1 = ((OptionSetValue)auditRecord.Attributes["operation"]).Value;
                            string str10 = num1.ToString();
                            str8 = str9 + "," + str10 + "," + auditRecord.FormattedValues["operation"];
                        }
                        else
                            str8 = str7 + "," + string.Empty + "," + string.Empty;
                        string str11;
                        if (auditRecord.Attributes.Contains("userid"))
                        {
                            EntityReference attributeValue = auditRecord.GetAttributeValue<EntityReference>("userid");
                            str11 = str8 + "," + attributeValue.Id.ToString() + "," + attributeValue.Name.ToString();
                        }
                        else
                            str11 = str8 + "," + string.Empty + "," + string.Empty;
                        string str12 = str11 + ",";
                        string empty4 = string.Empty;
                        if (auditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                        {
                            Dictionary<string, object> dictionary1 = new Dictionary<string, object>();
                            Dictionary<string, object> dictionary2 = new Dictionary<string, object>();
                            Entity newValue = ((AttributeAuditDetail)auditDetail).NewValue;
                            if (newValue != null)
                                dictionary2 = newValue.Attributes.ToDictionary<KeyValuePair<string, object>, string, object>((Func<KeyValuePair<string, object>, string>)(v => v.Key), (Func<KeyValuePair<string, object>, object>)(v => v.Value));
                            Entity oldValue = ((AttributeAuditDetail)auditDetail).OldValue;
                            if (oldValue != null)
                                dictionary1 = oldValue.Attributes.ToDictionary<KeyValuePair<string, object>, string, object>((Func<KeyValuePair<string, object>, string>)(v => v.Key), (Func<KeyValuePair<string, object>, object>)(v => v.Value));
                            if (dictionary2.Count > 0)
                            {
                                foreach (KeyValuePair<string, object> keyValuePair in dictionary2)
                                {
                                    KeyValuePair<string, object> de = keyValuePair;
                                    string str13 = lstEntityAttribute.FirstOrDefault<EntityAttribute>((Func<EntityAttribute, bool>)(o => o.LogicalName == de.Key.ToString())).DisplayName + ",";
                                    string str14;
                                    if (dictionary1.ContainsKey(de.Key))
                                    {
                                        object obj = dictionary1[de.Key.ToString()];
                                        if (obj.GetType() == typeof(Decimal))
                                            str14 = str13 + string.Empty + "," + obj.ToString() + "," + string.Empty + ",";
                                        else if (obj.GetType() == typeof(double))
                                            str14 = str13 + string.Empty + "," + obj.ToString() + "," + string.Empty + ",";
                                        else if (obj.GetType() == typeof(DateTime))
                                        {
                                            ((DateTime)obj).AddMinutes(330.0);
                                            str14 = str13 + string.Empty + "," + obj.ToString() + "," + string.Empty + ",";
                                        }
                                        else if (obj.GetType() == typeof(EntityReference))
                                        {
                                            EntityReference entityReference = (EntityReference)obj;
                                            if (entityReference != null)
                                                str14 = str13 + entityReference.Id.ToString() + "," + Program.RemoveSpecialCharacters(entityReference.Name == null ? "" : entityReference.Name.ToString()) + "," + entityReference.LogicalName + ",";
                                            else
                                                str14 = str13 + string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                        }
                                        else if (obj.GetType() == typeof(bool))
                                        {
                                            string boolText = Program.GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(obj));
                                            str14 = str13 + obj.ToString() + "," + boolText + "," + string.Empty + ",";
                                        }
                                        else if (obj.GetType() == typeof(int))
                                            str14 = str13 + string.Empty + "," + obj.ToString() + "," + string.Empty + ",";
                                        else if (obj.GetType() == typeof(Money))
                                        {
                                            Money money = (Money)obj;
                                            str14 = str13 + string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                        }
                                        else if (obj.GetType() == typeof(OptionSetValue))
                                        {
                                            OptionSetValue optValue = (OptionSetValue)obj;
                                            EntityOptionSet entityOptionSet = lstEntityAttribute.FirstOrDefault<EntityAttribute>((Func<EntityAttribute, bool>)(o => o.LogicalName == de.Key.ToString())).Optionset.FirstOrDefault<EntityOptionSet>((Func<EntityOptionSet, bool>)(o =>
                                            {
                                                int? optionValue = o.OptionValue;
                                                int num2 = optValue.Value;
                                                return optionValue.GetValueOrDefault() == num2 & optionValue.HasValue;
                                            }));
                                            if (entityOptionSet != null)
                                            {
                                                string[] strArray3 = new string[7];
                                                strArray3[0] = str13;
                                                num1 = optValue.Value;
                                                strArray3[1] = num1.ToString();
                                                strArray3[2] = ",";
                                                strArray3[3] = Program.RemoveSpecialCharacters(entityOptionSet.OptionName);
                                                strArray3[4] = ",";
                                                strArray3[5] = string.Empty;
                                                strArray3[6] = ",";
                                                str14 = string.Concat(strArray3);
                                            }
                                            else
                                            {
                                                string[] strArray4 = new string[5]
                                                {
                          str13,
                          null,
                          null,
                          null,
                          null
                                                };
                                                num1 = optValue.Value;
                                                strArray4[1] = num1.ToString();
                                                strArray4[2] = ",,";
                                                strArray4[3] = string.Empty;
                                                strArray4[4] = ",";
                                                str14 = string.Concat(strArray4);
                                            }
                                        }
                                        else if (obj.GetType() == typeof(string))
                                            str14 = str13 + string.Empty + "," + Program.RemoveSpecialCharacters(obj.ToString()) + "," + string.Empty + ",";
                                        else
                                            str14 = str13 + string.Empty + "," + Program.RemoveSpecialCharacters(obj.ToString()) + "," + string.Empty + ",";
                                        dictionary1.Remove(de.Key.ToString());
                                    }
                                    else
                                        str14 = str13 + string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                    string str15;
                                    if (de.Value.GetType() == typeof(Decimal))
                                        str15 = str14 + string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                    else if (de.Value.GetType() == typeof(double))
                                        str15 = str14 + string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                    else if (de.Value.GetType() == typeof(DateTime))
                                    {
                                        DateTime dateTime = ((DateTime)de.Value).AddMinutes(330.0);
                                        str15 = str14 + string.Empty + "," + dateTime.ToString() + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(EntityReference))
                                    {
                                        EntityReference entityReference = (EntityReference)de.Value;
                                        if (entityReference != null)
                                            str15 = str14 + entityReference.Id.ToString() + "," + Program.RemoveSpecialCharacters(entityReference.Name == null ? "" : entityReference.Name.ToString()) + "," + entityReference.LogicalName + ",";
                                        else
                                            str15 = str14 + string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(bool))
                                    {
                                        string boolText = Program.GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                        str15 = str14 + de.Value.ToString() + "," + boolText + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(int))
                                        str15 = str14 + string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                    else if (de.Value.GetType() == typeof(Money))
                                    {
                                        Money money = (Money)de.Value;
                                        str15 = str14 + string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(OptionSetValue))
                                    {
                                        OptionSetValue optValue = (OptionSetValue)de.Value;
                                        EntityOptionSet entityOptionSet = lstEntityAttribute.FirstOrDefault<EntityAttribute>((Func<EntityAttribute, bool>)(o => o.LogicalName == de.Key.ToString())).Optionset.FirstOrDefault<EntityOptionSet>((Func<EntityOptionSet, bool>)(o =>
                                        {
                                            int? optionValue = o.OptionValue;
                                            int num3 = optValue.Value;
                                            return optionValue.GetValueOrDefault() == num3 & optionValue.HasValue;
                                        }));
                                        if (entityOptionSet != null)
                                        {
                                            string[] strArray5 = new string[7];
                                            strArray5[0] = str14;
                                            num1 = optValue.Value;
                                            strArray5[1] = num1.ToString();
                                            strArray5[2] = ",";
                                            strArray5[3] = Program.RemoveSpecialCharacters(entityOptionSet.OptionName);
                                            strArray5[4] = ",";
                                            strArray5[5] = string.Empty;
                                            strArray5[6] = ",";
                                            str15 = string.Concat(strArray5);
                                        }
                                        else
                                        {
                                            string[] strArray6 = new string[5]
                                            {
                        str14,
                        null,
                        null,
                        null,
                        null
                                            };
                                            num1 = optValue.Value;
                                            strArray6[1] = num1.ToString();
                                            strArray6[2] = ",,";
                                            strArray6[3] = string.Empty;
                                            strArray6[4] = ",";
                                            str15 = string.Concat(strArray6);
                                        }
                                    }
                                    else if (de.Value.GetType() == typeof(string))
                                        str15 = str14 + string.Empty + "," + Program.RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                    else
                                        str15 = str14 + string.Empty + "," + Program.RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                    if (str15.Length > 0)
                                        str15 = str15.Substring(0, str15.Length - 1);
                                    stringBuilder.AppendLine(str12 + str15);
                                }
                            }
                            if (dictionary1.Count > 0)
                            {
                                foreach (KeyValuePair<string, object> keyValuePair in dictionary1)
                                {
                                    KeyValuePair<string, object> de = keyValuePair;
                                    string str16 = lstEntityAttribute.FirstOrDefault<EntityAttribute>((Func<EntityAttribute, bool>)(o => o.LogicalName == de.Key.ToString())).DisplayName + ",";
                                    string str17;
                                    if (de.Value.GetType() == typeof(Decimal))
                                        str17 = str16 + string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                    else if (de.Value.GetType() == typeof(double))
                                        str17 = str16 + string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                    else if (de.Value.GetType() == typeof(DateTime))
                                    {
                                        DateTime dateTime = ((DateTime)de.Value).AddMinutes(330.0);
                                        str17 = str16 + string.Empty + "," + dateTime.ToString() + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(EntityReference))
                                    {
                                        EntityReference entityReference = (EntityReference)de.Value;
                                        if (entityReference != null)
                                            str17 = str16 + entityReference.Id.ToString() + "," + Program.RemoveSpecialCharacters(entityReference.Name == null ? "" : entityReference.Name.ToString()) + "," + entityReference.LogicalName + ",";
                                        else
                                            str17 = str16 + string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(bool))
                                    {
                                        string boolText = Program.GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                        str17 = str16 + de.Value.ToString() + "," + boolText + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(int))
                                        str17 = str16 + string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                    else if (de.Value.GetType() == typeof(Money))
                                    {
                                        Money money = (Money)de.Value;
                                        str17 = str16 + string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                    }
                                    else if (de.Value.GetType() == typeof(OptionSetValue))
                                    {
                                        OptionSetValue optValue = (OptionSetValue)de.Value;
                                        EntityOptionSet entityOptionSet = lstEntityAttribute.FirstOrDefault<EntityAttribute>((Func<EntityAttribute, bool>)(o => o.LogicalName == de.Key.ToString())).Optionset.FirstOrDefault<EntityOptionSet>((Func<EntityOptionSet, bool>)(o =>
                                        {
                                            int? optionValue = o.OptionValue;
                                            int num4 = optValue.Value;
                                            return optionValue.GetValueOrDefault() == num4 & optionValue.HasValue;
                                        }));
                                        if (entityOptionSet != null)
                                        {
                                            string[] strArray7 = new string[7];
                                            strArray7[0] = str16;
                                            num1 = optValue.Value;
                                            strArray7[1] = num1.ToString();
                                            strArray7[2] = ",";
                                            strArray7[3] = Program.RemoveSpecialCharacters(entityOptionSet.OptionName);
                                            strArray7[4] = ",";
                                            strArray7[5] = string.Empty;
                                            strArray7[6] = ",";
                                            str17 = string.Concat(strArray7);
                                        }
                                        else
                                        {
                                            string[] strArray8 = new string[5]
                                            {
                        str16,
                        null,
                        null,
                        null,
                        null
                                            };
                                            num1 = optValue.Value;
                                            strArray8[1] = num1.ToString();
                                            strArray8[2] = ",,";
                                            strArray8[3] = string.Empty;
                                            strArray8[4] = ",";
                                            str17 = string.Concat(strArray8);
                                        }
                                    }
                                    else if (de.Value.GetType() == typeof(string))
                                        str17 = str16 + string.Empty + "," + Program.RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                    else
                                        str17 = str16 + string.Empty + "," + Program.RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                    string str18 = str17 + string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                    if (str18.Length > 0)
                                        str18 = str18.Substring(0, str18.Length - 1);
                                    stringBuilder.AppendLine(str12 + str18);
                                }
                            }
                        }
                        else if (auditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                        {
                            UserAccessAuditDetail accessAuditDetail = (UserAccessAuditDetail)auditDetail;
                            string[] strArray9 = new string[14]
                            {
                ",",
                string.Empty,
                ",",
                string.Empty,
                ",",
                string.Empty,
                ",",
                string.Empty,
                ",",
                null,
                null,
                null,
                null,
                null
                            };
                            DateTime dateTime = accessAuditDetail.AccessTime;
                            dateTime = dateTime.AddMinutes(330.0);
                            strArray9[9] = dateTime.ToString();
                            strArray9[10] = ",";
                            strArray9[11] = string.Empty;
                            strArray9[12] = ",";
                            strArray9[13] = string.Empty;
                            string str19 = string.Concat(strArray9);
                            stringBuilder.AppendLine(str12 + str19);
                        }
                        else
                        {
                            string str20 = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                            stringBuilder.AppendLine(str12 + str20);
                        }
                        flag = true;
                    }
                    else
                    {
                        Console.WriteLine(formattedValue);
                        DeleteRecordChangeHistoryRequest request2 = new DeleteRecordChangeHistoryRequest();
                        request2.Target = new EntityReference(entityLogicalName, entityId);
                        try
                        {
                            _service.Execute((OrganizationRequest)request2);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("error on Audit Log Delition " + ex.Message);
                        }
                    }
                }
                if (stringBuilder.ToString().Length > 0 & flag)
                {
                    Program.UploadAzureBlob(entityId.ToString(), stringBuilder.ToString(), container);
                    DeleteRecordChangeHistoryRequest request3 = new DeleteRecordChangeHistoryRequest();
                    request3.Target = new EntityReference(entityLogicalName, entityId);
                    try
                    {
                        _service.Execute((OrganizationRequest)request3);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error on Audit Log Delition " + ex.Message);
                    }
                }
                else
                {
                    DeleteRecordChangeHistoryRequest request4 = new DeleteRecordChangeHistoryRequest();
                    request4.Target = new EntityReference(entityLogicalName, entityId);
                    try
                    {
                        _service.Execute((OrganizationRequest)request4);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error on Audit Log Delition " + ex.Message);
                    }
                    Console.WriteLine(" ##Error Not deleted ");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(" ##Error " + ex.Message);
            }
        }
        static void updateFlag(string entityLogicalName, string _fromDate, string _toDate, IOrganizationService _service)
        {
            try
            {

                QueryExpression query = new QueryExpression(entityLogicalName);
                query.ColumnSet = new ColumnSet("hil_isauditlogmigrated");
                query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _fromDate);
                query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, _toDate);
                query.AddOrder("createdon", OrderType.Ascending);
                query.NoLock = true;
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                EntityCollection ec = _service.RetrieveMultiple(query);
                do
                {
                    foreach (Entity entity1 in ec.Entities)
                    {
                        try
                        {
                            Entity entity = new Entity(entityLogicalName, entity1.Id);
                            entity["hil_isauditlogmigrated"] = false;
                            _service.Update(entity);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error in updateingFlag " + ex.Message);
                        }
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(query);
                }
                while (ec.MoreRecords);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in UpdateFlag " + ex.Message);
            }
        }
        public static IOrganizationService createConnection(string connectionString)
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
        public static void GetAuditEnableEntity(IOrganizationService service)
        {
            RetrieveAllEntitiesRequest metaDataRequest = new RetrieveAllEntitiesRequest();
            RetrieveAllEntitiesResponse metaDataResponse = new RetrieveAllEntitiesResponse();
            metaDataRequest.EntityFilters = EntityFilters.Entity;

            XmlDictionaryReaderQuotas myReaderQuotas = new XmlDictionaryReaderQuotas();
            myReaderQuotas.MaxNameTableCharCount = 2147483647;

            // Execute the request.

            metaDataResponse = (RetrieveAllEntitiesResponse)service.Execute(metaDataRequest);

            var entities = metaDataResponse.EntityMetadata;
            foreach (var abc in entities)
            {
                //if (abc.ObjectTypeCode == 10460||abc.ObjectTypeCode== 10695)
                //if (abc.ObjectTypeCode != 10196)
                //    continue;
                int abccc = Array.Find(notLoggedentity, element => element == abc.ObjectTypeCode);
                if (abc.IsAuditEnabled.Value == true && abccc == 0)
                {
                    string key = abc.EntitySetName;
                    String logicalName = abc.LogicalName;
                    string schemaName = abc.SchemaName;
                    string primaryFieldName = abc.PrimaryNameAttribute;
                    string primaryIDName = abc.PrimaryIdAttribute;
                    //AuditLogEnabledEntityList.Add(key, abc.ObjectTypeCode);
                    //if(abc.ObjectTypeCode == 10681)
                    //{
                    //    Console.WriteLine("Entity Name " + key + " type " + abc.ObjectTypeCode);

                    //}
                    D365AuditLogMigration_Delta(logicalName, service, primaryFieldName, primaryIDName);
                }
            }
        }
        static void D365AuditLogMigration_DeltaWithFlag(string _entityName, string _fromDate, string _toDate, IOrganizationService _service)
        {
            int i = 1;
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            GetPrimaryIdFieldName(_entityName, out _primaryKey, out _primaryfield, _service);
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes(_entityName, _service);
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = _entityName;
            string[] primaryKeySchemaName = { _primaryKey, _primaryfield, "createdon" };

            QueryExpression query = new QueryExpression(entityLogicalName);
            query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            query.Criteria.AddCondition("hil_isauditlogmigrated", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _fromDate);
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, _toDate);
            query.AddOrder("createdon", OrderType.Ascending);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        string _recID = string.Empty;

                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (ent.Attributes.Contains(_primaryfield))
                            {
                                _recID = ent.GetAttributeValue<string>(_primaryfield);
                            }
                            else { _recID = ""; }

                            StringBuilder strContent = new StringBuilder();
                            if (!CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            }
                            Console.WriteLine("Record # " + i.ToString() + " " + _recID + " Start..." + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            changeRequest.Target.LogicalName = entityLogicalName;
                            RetrieveRecordChangeHistoryResponse changeResponse = null;

                        Repeat:
                            try
                            {
                                changeResponse = (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);
                            }
                            catch
                            {
                                goto Repeat;
                            }

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            string strLine;
                            string strAttributes;
                            string dateString;
                            bool _logExists = false;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                string _action = auditRecord.FormattedValues["action"];

                                if (auditRecord.Attributes.Contains("objectid") && (_action == "Create" || _action == "Update" || _action == "Delete"))
                                {
                                    if (auditRecord.Attributes.Contains("auditid"))
                                    {
                                        strLine = auditRecord.Attributes["auditid"].ToString();
                                    }
                                    else
                                    {
                                        strLine = null;
                                    }
                                    if (auditRecord.Attributes.Contains("createdon"))
                                    {
                                        DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                        strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("action"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["action"];
                                    }
                                    else
                                    {

                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("objectid"))
                                    {
                                        strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("objecttypecode"))
                                    {
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("operation"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["operation"];
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("userid"))
                                    {
                                        EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                        strLine += "," + er.Id.ToString();
                                        strLine += "," + er.Name.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    strLine += ",";
                                    strAttributes = string.Empty;
                                    if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                    {
                                        Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                        Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                        object _OldValue;
                                        //object _NewValue;
                                        var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                        if (newValueEntity != null)
                                        {
                                            newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }

                                        var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                        if (oldValueEntity != null)
                                        {
                                            oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }
                                        if (newValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in newValues)
                                            {
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                //strAttributes = de.Key.ToString() + ",";
                                                //Console.WriteLine("Attribute Name :" + strAttributes);
                                                if (oldValues.ContainsKey(de.Key))
                                                {
                                                    _OldValue = oldValues[de.Key.ToString()];
                                                    if (_OldValue.GetType() == typeof(Decimal))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Double))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.DateTime))
                                                    {
                                                        DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                    {
                                                        EntityReference entRef = (EntityReference)_OldValue;
                                                        if (entRef != null)
                                                        {
                                                            strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                        }
                                                        else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Boolean))
                                                    {
                                                        string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(_OldValue));
                                                        //string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                        strAttributes += _OldValue.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Int32))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                    {
                                                        Money money = (Money)_OldValue;
                                                        strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                    {
                                                        OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                        EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                        EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                        if (optionSet != null)
                                                        {
                                                            strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                        }
                                                        else
                                                        {
                                                            strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                        }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.String))
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }

                                                    oldValues.Remove(de.Key.ToString());
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                }

                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }

                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                        if (oldValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in oldValues)
                                            {
                                                //strAttributes = de.Key.ToString() + ",";
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                    }
                                    else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                    {
                                        var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    else
                                    {
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    _logExists = true;
                                }
                                else
                                {
                                    Console.WriteLine(_action);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            if (strContent.ToString().Length > 0 && _logExists)
                            {
                                UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                                var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                                try
                                {
                                    _service.Execute(DeleteAudit);
                                }
                                catch { }
                            }
                            entTemp = new Entity(entityLogicalName, ent.Id);
                            entTemp["hil_isauditlogmigrated"] = true;
                            _service.Update(entTemp);
                            Console.WriteLine("Record # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }

        static void D365AuditLogMigration_Delta(string _entityName, IOrganizationService _service, string _primaryfield, string _primaryKey)//(string _entityName, string _fromDate, string _toDate, IOrganizationService _service)
        {
            string _fromDate = "2022-08-01";// DateTime.Today.Year.ToString() + "-" + DateTime.Today.Month.ToString().PadLeft(2,'0') + "-" + (DateTime.Today.Day-1).ToString().PadLeft(2, '0');
            Console.WriteLine("###########################################################################");
            Console.WriteLine("Entity Name " + _entityName);
            int i = 1;
            //string _primaryfield = string.Empty;
            //string _primaryKey = string.Empty;
            //  GetPrimaryIdFieldName(_entityName, out _primaryKey, out _primaryfield, _service);
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes(_entityName, _service);
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = _entityName;
            string[] primaryKeySchemaName = { _primaryKey, _primaryfield, "createdon" };

            QueryExpression query = new QueryExpression(entityLogicalName);
            query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            //query.Criteria.AddCondition("hil_isauditlogmigrated", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _fromDate);
            //query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, _toDate);
            query.AddOrder("createdon", OrderType.Ascending);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        string _recID = string.Empty;

                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (ent.Attributes.Contains(_primaryfield))
                            {
                                _recID = ent.GetAttributeValue<string>(_primaryfield);
                            }
                            else { _recID = ""; }

                            StringBuilder strContent = new StringBuilder();
                            if (!CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            }
                            Console.WriteLine("Record # " + i.ToString() + " " + _recID + " Start..." + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse = null;

                        Repeat:
                            try
                            {
                                changeResponse = (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);
                            }
                            catch
                            {
                                goto Repeat;
                            }

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            string strLine;
                            string strAttributes;
                            string dateString;
                            bool _logExists = false;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                string _action = auditRecord.FormattedValues["action"];

                                if (auditRecord.Attributes.Contains("objectid") && (_action == "Create" || _action == "Update" || _action == "Delete"))
                                {
                                    if (auditRecord.Attributes.Contains("auditid"))
                                    {
                                        strLine = auditRecord.Attributes["auditid"].ToString();
                                    }
                                    else
                                    {
                                        strLine = null;
                                    }
                                    if (auditRecord.Attributes.Contains("createdon"))
                                    {
                                        DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                        strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("action"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["action"];
                                    }
                                    else
                                    {

                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("objectid"))
                                    {
                                        strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("objecttypecode"))
                                    {
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("operation"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["operation"];
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("userid"))
                                    {
                                        EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                        strLine += "," + er.Id.ToString();
                                        strLine += "," + er.Name.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    strLine += ",";
                                    strAttributes = string.Empty;
                                    if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                    {
                                        Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                        Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                        object _OldValue;
                                        //object _NewValue;
                                        var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                        if (newValueEntity != null)
                                        {
                                            newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }

                                        var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                        if (oldValueEntity != null)
                                        {
                                            oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }
                                        if (newValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in newValues)
                                            {
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                //strAttributes = de.Key.ToString() + ",";
                                                //Console.WriteLine("Attribute Name :" + strAttributes);
                                                if (oldValues.ContainsKey(de.Key))
                                                {
                                                    _OldValue = oldValues[de.Key.ToString()];
                                                    if (_OldValue.GetType() == typeof(Decimal))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Double))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.DateTime))
                                                    {
                                                        DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                    {
                                                        EntityReference entRef = (EntityReference)_OldValue;
                                                        if (entRef != null)
                                                        {
                                                            strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                        }
                                                        else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Boolean))
                                                    {
                                                        string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(_OldValue));
                                                        //string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                        strAttributes += _OldValue.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Int32))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                    {
                                                        Money money = (Money)_OldValue;
                                                        strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                    {
                                                        OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                        EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                        EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                        if (optionSet != null)
                                                        {
                                                            strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                        }
                                                        else
                                                        {
                                                            strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                        }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.String))
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }

                                                    oldValues.Remove(de.Key.ToString());
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                }

                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }

                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                        if (oldValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in oldValues)
                                            {
                                                //strAttributes = de.Key.ToString() + ",";
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                    }
                                    else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                    {
                                        var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    else
                                    {
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    _logExists = true;
                                }
                                else
                                {
                                    Console.WriteLine(_action);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            if (strContent.ToString().Length > 0 && _logExists)
                            {
                                UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                                var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                                try
                                {
                                    _service.Execute(DeleteAudit);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Error " + ex.Message);
                                }
                            }
                            entTemp = new Entity(entityLogicalName, ent.Id);
                            entTemp["hil_isauditlogmigrated"] = true;
                            _service.Update(entTemp);
                            Console.WriteLine("Record # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        #region Common Libraries
        static List<EntityAttribute> GetEntityAttributes(string _entityLogicalName, IOrganizationService _service)
        {
            List<EntityAttribute> lstAttribute = new List<EntityAttribute>();

            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = _entityLogicalName,
                RetrieveAsIfPublished = true
            };
            RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)_service.Execute(retrieveEntityRequest);
            EntityMetadata AccountEntity = retrieveAccountEntityResponse.EntityMetadata;

            foreach (object attribute in AccountEntity.Attributes)
            {
                AttributeMetadata a = (AttributeMetadata)attribute;
                if (a.AttributeTypeName.Value != "VirtualType")
                {
                    EntityAttribute entAttr = new EntityAttribute();
                    entAttr.LogicalName = a.LogicalName;
                    if (a.DisplayName.LocalizedLabels.Count > 0)
                    { entAttr.DisplayName = a.DisplayName.LocalizedLabels[0].Label; }
                    else
                    { entAttr.DisplayName = a.LogicalName; }
                    entAttr.AttributeType = a.AttributeTypeName.Value.ToString();
                    if (entAttr.AttributeType.IndexOf("Pick") >= 0 || entAttr.AttributeType.IndexOf("State") >= 0 || entAttr.AttributeType.IndexOf("Status") >= 0)
                    {
                        entAttr.Optionset = new List<EntityOptionSet>();
                        var retrieveAttributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = _entityLogicalName,
                            LogicalName = entAttr.LogicalName
                        };
                        var retrieveAttributeResponse = (RetrieveAttributeResponse)_service.Execute(retrieveAttributeRequest);
                        var retrievedPicklistAttributeMetadata = (EnumAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                        foreach (OptionMetadata opt in retrievedPicklistAttributeMetadata.OptionSet.Options)
                        {
                            entAttr.Optionset.Add(new EntityOptionSet() { OptionName = opt.Label.LocalizedLabels[0].Label, OptionValue = opt.Value });
                        }
                    }
                    lstAttribute.Add(entAttr);
                }
            }
            return lstAttribute;
        }
        static void GetPrimaryIdFieldName(string _entityName, out string _primaryKey, out string _primaryField, IOrganizationService _service)
        {
            //Create RetrieveEntityRequest
            _primaryKey = "";
            _primaryField = "";
            try
            {
                RetrieveEntityRequest retrievesEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = _entityName
                };

                //Execute Request
                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)_service.Execute(retrievesEntityRequest);
                _primaryKey = retrieveEntityResponse.EntityMetadata.PrimaryIdAttribute;

                _primaryField = retrieveEntityResponse.EntityMetadata.PrimaryNameAttribute;
            }
            catch (Exception ex)
            {
                _primaryKey = "ERROR";
                _primaryKey = ex.Message;
            }
        }
        static string RemoveSpecialCharacters(string _attributeValue)
        {
            return _attributeValue.Replace(", ", " ").Replace(",", " ").Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("@", "at");
        }
        static string GetBoolText(IOrganizationService service, string entitySchemaName, string attributeSchemaName, bool value)
        {
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entitySchemaName,
                LogicalName = attributeSchemaName,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            BooleanAttributeMetadata retrievedBooleanAttributeMetadata = (BooleanAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            string boolText = string.Empty;
            if (value)
            {
                boolText = retrievedBooleanAttributeMetadata.OptionSet.TrueOption.Label.UserLocalizedLabel.Label;
            }
            else
            {
                boolText = retrievedBooleanAttributeMetadata.OptionSet.FalseOption.Label.UserLocalizedLabel.Label;
            }
            return boolText;
        }
        static void UploadAzureBlob(string _fileName, string _strContent, CloudBlobContainer container)
        {
            // This creates a reference to the append blob.  
            CloudAppendBlob appBlob = container.GetAppendBlobReference(_fileName + ".csv");

            // Now we are going to check if todays file exists and if it doesn't we create it.  
            if (!appBlob.Exists())
            {
                appBlob.CreateOrReplace();
            }

            // Add the entry to file.  
            _strContent = _strContent.Replace("@", System.Environment.NewLine);
            appBlob.AppendText
            (
            string.Format(
                    "{0}\r\n",
                    _strContent)
             );
        }
        static bool CheckIfAzureBlobExist(string _fileName, CloudBlobContainer container)
        {
            bool _retValue = false;
            CloudAppendBlob appBlob = container.GetAppendBlobReference(_fileName + ".csv");

            if (appBlob.Exists())
            {
                _retValue = true;
            }
            return _retValue;
        }
        #endregion
    }
    public class EntityAttribute
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string AttributeType { get; set; }
        public List<EntityOptionSet> Optionset { get; set; }
    }
    public class EntityOptionSet
    {
        public string OptionName { get; set; }
        public int? OptionValue { get; set; }
    }
}
