
using Microsoft.Crm.Sdk.Messages;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using Microsoft.Xrm.Sdk.Client;
using System.Configuration;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Description;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class D365AuditLog
    {
        [DataMember]
        public string EntityName { get; set; }
        [DataMember]
        public string RecId { get; set; }
        [DataMember]
        public string FromDate { get; set; }
        [DataMember]
        public string ToDate { get; set; }
        [DataMember]
        public string FieldName { get; set; }

        public List<AttributesMetadata> GetAttributeMetadata(string _entityLogicalName)
        {
            List<AttributesMetadata> lstAttributeMetadata;
            try
            {
                IOrganizationService _service = ConnectToCRM.GetOrgService();
                if (_service != null)
                {
                    if (_entityLogicalName.Trim().Length == 0 || _entityLogicalName == null)
                    {
                        lstAttributeMetadata = new List<AttributesMetadata>() { new AttributesMetadata() { ResultStatus = false, ResultMessage = "Entity Logical name is required." } };
                        return lstAttributeMetadata;
                    }

                    List<AttributesMetadata> lstAttribute = new List<AttributesMetadata>();

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
                            AttributesMetadata entAttr = new AttributesMetadata();
                            entAttr.LogicalName = a.LogicalName;
                            if (a.DisplayName.LocalizedLabels.Count > 0)
                            { entAttr.DisplayName = a.DisplayName.LocalizedLabels[0].Label; }
                            else
                            { entAttr.DisplayName = a.LogicalName; }
                            entAttr.AttributeType = a.AttributeTypeName.Value.ToString();
                            lstAttribute.Add(entAttr);
                        }
                    }
                    return lstAttribute;
                }
                else
                {
                    lstAttributeMetadata = new List<AttributesMetadata>() { new AttributesMetadata() { ResultStatus = false, ResultMessage = "D365 Service Unavailable." } };
                    return lstAttributeMetadata;
                }
            }
            catch (Exception ex)
            {
                lstAttributeMetadata = new List<AttributesMetadata>() { new AttributesMetadata() { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message } };
                return lstAttributeMetadata;
            }
        }
        public List<D365AuditLogResult> GetD365AuditLogData(D365AuditLog _requestData)
        {
            List<D365AuditLogResult> objD365AuditLog = new List<D365AuditLogResult>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgServiceProd(); //get org service obj for connection
                D365AuditLogMigrationDaily(_requestData.EntityName, _requestData.RecId, service);

                CloudAppendBlob _apprndBlob = ConnectWithAzureBlob(_requestData.RecId);
                DateTime _fromDate, _toDate;

                if (_requestData.FromDate == null || _requestData.FromDate.Length == 0)
                {
                    _fromDate = new DateTime(1900, 1, 1);
                }
                else
                {
                    _fromDate = ConvertStringToDate(_requestData.FromDate);
                }

                if (_requestData.ToDate == null || _requestData.ToDate.Length == 0)
                {
                    _toDate = new DateTime(2099, 1, 1);
                }
                else
                {
                    _toDate = ConvertStringToDate(_requestData.ToDate);
                }

                if (_apprndBlob == null)
                {
                    objD365AuditLog.Add(new D365AuditLogResult() { Status = "204", StatusRemarks = "ERROR !!! Something went wrong." });
                    return objD365AuditLog;
                }

                string _blobContent = string.Empty;
                using (var memoryStream = new MemoryStream())
                {
                    _apprndBlob.DownloadToStream(memoryStream);
                    _blobContent = System.Text.Encoding.UTF8.GetString(memoryStream.ToArray());
                }
                string[] _blobContentLines = _blobContent.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                string[] _blobLineFields = null;
                DateTime _tempDate;
                foreach (string _line in _blobContentLines)
                {
                    _blobLineFields = _line.Split(',');
                    if (_blobLineFields.Length >= 18 && _blobLineFields[0] != "auditid")
                    {
                        _tempDate = ConvertStringToDate(_blobLineFields[1]);
                        if (_tempDate >= _fromDate && _tempDate <= _toDate)
                        {
                            string _action = Convert.ToString(_blobLineFields[3]);
                            if (_requestData.FieldName == "All" || string.IsNullOrEmpty(_requestData.FieldName) && (_action == "Create" || _action == "Update" || _action == "Delete"))
                            {
                                objD365AuditLog.Add(new D365AuditLogResult
                                {
                                    AuditId = Convert.ToString(_blobLineFields[0]),
                                    CreatedOnDt = ConvertStringToDate(_blobLineFields[1]),
                                    CreatedOn = _blobLineFields[1],
                                    Action = Convert.ToInt32(_blobLineFields[2]),
                                    ActionName = Convert.ToString(_blobLineFields[3]),
                                    ObjectId = new Guid(_blobLineFields[4]),
                                    ObjectType = Convert.ToString(_blobLineFields[5]),
                                    Operation = Convert.ToInt32(_blobLineFields[7]),
                                    OperationName = Convert.ToString(_blobLineFields[8]),
                                    UserId = new Guid(_blobLineFields[9]),
                                    UserIdName = Convert.ToString(_blobLineFields[10]),
                                    FieldName = Convert.ToString(_blobLineFields[11]),
                                    OldValue = Convert.ToString(_blobLineFields[12]),
                                    OldValueName = Convert.ToString(_blobLineFields[13]),
                                    OldValueType = Convert.ToString(_blobLineFields[14]),
                                    NewValue = Convert.ToString(_blobLineFields[15]),
                                    NewValueName = Convert.ToString(_blobLineFields[16]),
                                    NewValueType = Convert.ToString(_blobLineFields[17])
                                }); ;
                            }
                            else
                            {
                                if (_requestData.FieldName.ToUpper() == Convert.ToString(_blobLineFields[11]).ToUpper())
                                {
                                    objD365AuditLog.Add(new D365AuditLogResult
                                    {
                                        AuditId = Convert.ToString(_blobLineFields[0]),
                                        CreatedOnDt = ConvertStringToDate(_blobLineFields[1]),
                                        CreatedOn = _blobLineFields[1],
                                        Action = Convert.ToInt32(_blobLineFields[2]),
                                        ActionName = Convert.ToString(_blobLineFields[3]),
                                        ObjectId = new Guid(_blobLineFields[4]),
                                        ObjectType = Convert.ToString(_blobLineFields[5]),
                                        Operation = Convert.ToInt32(_blobLineFields[7]),
                                        OperationName = Convert.ToString(_blobLineFields[8]),
                                        UserId = new Guid(_blobLineFields[9]),
                                        UserIdName = Convert.ToString(_blobLineFields[10]),
                                        FieldName = Convert.ToString(_blobLineFields[11]),
                                        OldValue = Convert.ToString(_blobLineFields[12]),
                                        OldValueName = Convert.ToString(_blobLineFields[13]),
                                        OldValueType = Convert.ToString(_blobLineFields[14]),
                                        NewValue = Convert.ToString(_blobLineFields[15]),
                                        NewValueName = Convert.ToString(_blobLineFields[16]),
                                        NewValueType = Convert.ToString(_blobLineFields[17])
                                    });
                                }
                            }
                        }
                    }
                }
                objD365AuditLog = objD365AuditLog.OrderByDescending(p => p.CreatedOnDt).ToList();
                return objD365AuditLog;
            }
            catch (Exception ex)
            {
                objD365AuditLog.Add(new D365AuditLogResult() { Status = "204", StatusRemarks = "ERROR !!! " + ex.Message });
                return objD365AuditLog;
            }
        }

        private void D365AuditLogMigrationDaily(string _entityName, string _entityID, IOrganizationService _service)
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
            //query.Criteria.AddCondition("hil_isauditlogmigrated", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition(_primaryKey, ConditionOperator.Equal, new Guid(_entityID));
            query.NoLock = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(query);

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

                            if (auditRecord.Attributes.Contains("objectid") && (_action == "Create" || _action == "Update" || _action == "Delete" || _action == "Assign" || _action == "Activate" || _action == "Deactivate"))
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
                }
            }
            catch
            {
                // Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        private void GetPrimaryIdFieldName(string _entityName, out string _primaryKey, out string _primaryField, IOrganizationService _service)
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

        private List<EntityAttribute> GetEntityAttributes(string _entityLogicalName, IOrganizationService _service)
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

        private DateTime ConvertStringToDate(string _strDate)
        {
            int year = 0, month = 0, day = 0, hh = 0, mm = 0, ss = 0;
            year = Convert.ToInt32(_strDate.Substring(0, 4));
            month = Convert.ToInt32(_strDate.Substring(4, 2));
            day = Convert.ToInt32(_strDate.Substring(6, 2));
            hh = Convert.ToInt32(_strDate.Substring(8, 2));
            mm = Convert.ToInt32(_strDate.Substring(10, 2));
            ss = Convert.ToInt32(_strDate.Substring(12, 2));
            return new DateTime(year, month, day, hh, mm, ss);
        }
        private CloudAppendBlob ConnectWithAzureBlob(string _recGuId)
        {
            string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
            CloudAppendBlob appendBlobReference = null;
            try
            {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConnectionSting);
                // Create the blob client.
                CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
                // Retrieve reference to a previously created container.
                CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                // Retrieve reference to a blob named "test.csv"
                appendBlobReference = container.GetAppendBlobReference(string.Format("{0}.csv", _recGuId));
            }
            catch { }
            return appendBlobReference;
        }

        #region Common Libraries

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
    [DataContract]
    public class AttributesMetadata
    {
        [DataMember]
        public string LogicalName { get; set; }
        [DataMember]
        public string DisplayName { get; set; }
        [DataMember]
        public string AttributeType { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
    }
    [DataContract]
    public class EntityAttribute
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string AttributeType { get; set; }
        public List<EntityOptionSet> Optionset { get; set; }
    }
    [DataContract]
    public class EntityOptionSet
    {
        public string OptionName { get; set; }
        public int? OptionValue { get; set; }
    }
    public class D365AuditLogResult
    {
        [DataMember]
        public string AuditId { get; set; }
        [DataMember]
        public string CreatedOn { get; set; }

        [DataMember]
        public DateTime CreatedOnDt { get; set; }
        [DataMember]
        public int Action { get; set; }
        [DataMember]
        public string ActionName { get; set; }
        [DataMember]
        public Guid ObjectId { get; set; }
        [DataMember]
        public string ObjectType { get; set; }
        [DataMember]
        public int Operation { get; set; }
        [DataMember]
        public string OperationName { get; set; }
        [DataMember]
        public Guid UserId { get; set; }
        [DataMember]
        public string UserIdName { get; set; }
        [DataMember]
        public string FieldName { get; set; }
        [DataMember]
        public string OldValue { get; set; }
        [DataMember]
        public string OldValueName { get; set; }
        [DataMember]
        public string OldValueType { get; set; }
        [DataMember]
        public string NewValue { get; set; }
        [DataMember]
        public string NewValueName { get; set; }
        [DataMember]
        public string NewValueType { get; set; }
        [DataMember]
        public string Status { get; set; }
        [DataMember]
        public string StatusRemarks { get; set; }
    }


    public class APIExceptionLog
    {
        public static void LogAPIExecution(string msg, string FolderName = "")
        {
            #region AzureConnection
            string accName = "d365storagesa"; //This is your Azure Blob Storage Name
            string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            string _recID = string.Empty;

            CloudStorageAccount strAcc = new CloudStorageAccount(new StorageCredentials(accName, accKey), true);
            CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();

            CloudBlobContainer container = blobClient.GetContainerReference("d365errorlog");
            container.CreateIfNotExistsAsync();

            #endregion AzureConnection
            DateTime _now = DateTime.Now;
            string path = _now.ToString("ddMMMyyyy") + "/";

            if (FolderName != "")
            {
                path = FolderName + "/" + _now.ToString("ddMMMyyyy") + "/";
            }
            CloudAppendBlob appBlob = container.GetAppendBlobReference(path + DateTime.Today.ToString("ddMMMyyyy") + ".csv");
            if (!appBlob.Exists())
            {
                appBlob.CreateOrReplace();
                appBlob.AppendText
                (
                    string.Format("{0}\r\n", "Caller Name,API Name,Error,Request Data,Response Data,Start Datetime,End Datetime,Execution Time")
                 );
            }
            appBlob.AppendText
            (
                string.Format("{0}\r\n", msg)
             );
        }
    }
}
