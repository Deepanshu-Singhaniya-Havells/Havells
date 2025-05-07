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
            var fromDate = ConfigurationManager.AppSettings["fromDate"].ToString();
            var toDate = ConfigurationManager.AppSettings["toDate"].ToString();
            var objecttypeCode = ConfigurationManager.AppSettings["ObjectTypeCode"].ToString();

            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = createConnection(finalString);

            GetAuditLogHeader(service, int.Parse(objecttypeCode), fromDate, toDate);

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
                if (abc.ObjectTypeCode == 10460 || abc.ObjectTypeCode == 10695)
                    continue;
                //int abccc = Array.Find(notLoggedentity, element => element == abc.ObjectTypeCode);
                if (abc.IsAuditEnabled.Value == true)//&& abccc == 0)
                {
                    int? objecttypeCode = abc.ObjectTypeCode;
                    GetAuditLogHeader(service, objecttypeCode, fromDate, toDate);

                }
            }
        }
        public static void GetAuditLogHeader(IOrganizationService service, int? objectTypeCode, string fromDate, string toDate)
        {
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            List<EntityAttribute> lstEntityAttribute = new List<EntityAttribute>();
            Console.WriteLine("Record Migrating of Entity Code: " + objectTypeCode + " from: " + fromDate + " till: " + toDate);
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

            string fetchXml = $@"<fetch page='" + pageNo + $@"' version='1.0' mapping='logical' >
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
                                        <condition attribute='createdon' operator='gt' value='{fromDate}' />
                                        <condition attribute='createdon' operator='lt' value='{toDate}' />
                                   </filter>
                                 </entity>
                            </fetch>";
            EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));
            totalRecCount = entityCollection.Entities.Count + totalRecCount;

            Console.WriteLine("From " + fromDate + " To " + toDate + " Count " + entityCollection.Entities.Count);

            foreach (Entity entity in entityCollection.Entities)
            {
                Guid objId = entity.GetAttributeValue<EntityReference>("objectid").Id;
                Guid abccc = Array.Find(guids.ToArray(), element => element == objId);
                if (abccc == Guid.Empty && objId != Guid.Empty)
                {
                    guids.Add(objId);
                }
                if (entityLogicalName == string.Empty)
                {
                    entityLogicalName = entity.GetAttributeValue<EntityReference>("objectid").LogicalName;
                    GetPrimaryIdFieldName(entityLogicalName, out _primaryKey, out _primaryfield, service);
                    lstEntityAttribute = GetEntityAttributes(entityLogicalName, service);
                }
            }

            totalAuditRecCount = totalAuditRecCount + guids.Count;
            Console.WriteLine("Entity Name: " + entityLogicalName + " JOb Count is " + totalAuditRecCount);

            //Console.WriteLine("**************************************");
            Parallel.ForEach(guids, entityId =>
            {
                totalRecDone++;
                int i = totalRecDone;
                Console.WriteLine(" #################################### Record Started " + i + "/" + totalAuditRecCount + " #################################### ");
                Console.WriteLine(entityId);
                migrateAuditLogs(service, entityLogicalName, entityId, _primaryfield, _primaryKey, lstEntityAttribute, container);
                Console.WriteLine(" #################################### Record Completed " + i + "/" + totalAuditRecCount + " #################################### ");
                //Console.WriteLine("");
            });

            //foreach (Guid entityId in guids)
            //{
            //    totalRecDone++;
            //    Console.WriteLine(" #################################### Record Started " + totalRecDone + "/" + totalAuditRecCount + " #################################### ");
            //    Console.WriteLine(entityId);
            //    migrateAuditLogs(service, entityLogicalName, entityId, _primaryfield, _primaryKey, lstEntityAttribute, container);
            //    Console.WriteLine(" #################################### Record Completed " + totalRecDone + "/" + totalAuditRecCount + " #################################### ");
            //    Console.WriteLine(""); Console.WriteLine("");
            //}

            if (entityCollection.Entities.Count == 500)
            {
                pageNo++;
                //Console.WriteLine(" ######## Again Repete");
                goto repetRetrive;
            }
            Console.WriteLine("**************************************");
            Console.WriteLine("Record Migrated of Entity Name: " + entityLogicalName + " from " + fromDate + " till " + toDate);
        }
        static void migrateAuditLogs(IOrganizationService _service, string entityLogicalName, Guid entityId, string _primaryfield, string _primaryKey,
            List<EntityAttribute> lstEntityAttribute, CloudBlobContainer container)
        {
            //Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            string[] primaryKeySchemaName = { _primaryKey, _primaryfield, "createdon" };
            string _recID = string.Empty;
            try
            {
                _recID = entityId.ToString();// ent.GetAttributeValue<string>(_primaryfield);
                StringBuilder strContent = new StringBuilder();
                if (!CheckIfAzureBlobExist(entityId.ToString(), container))
                {
                    strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                }
                //Console.WriteLine("Record # " + _recID + " Start...");
                RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                changeRequest.Target = new EntityReference(entityLogicalName, entityId);
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

                    if (auditRecord.Attributes.Contains("objectid") && (_action == "Create" || _action == "Update" || _action == "Delete"
                        || _action == "Activate" || _action == "Deactivate" || _action == "Assign"))
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
                        var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                        DeleteAudit.Target = new EntityReference(entityLogicalName, entityId);// ent.ToEntityReference();
                        try
                        {
                            _service.Execute(DeleteAudit);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("error on Audit Log Delition " + ex.Message);
                        }
                    }
                }
                if (strContent.ToString().Length > 0 && _logExists)
                {
                    UploadAzureBlob(entityId.ToString(), strContent.ToString(), container);
                    #region Delete Audit Log
                    var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                    DeleteAudit.Target = new EntityReference(entityLogicalName, entityId);// ent.ToEntityReference();
                    try
                    {
                        _service.Execute(DeleteAudit);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error on Audit Log Delition " + ex.Message);
                    }

                    #endregion
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
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM."))
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
