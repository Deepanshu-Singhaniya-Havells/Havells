using System.Collections.Generic;
using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using System.Globalization;
using Microsoft.Crm.Sdk.Messages;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AttachmentManager
    {
        [DataMember]
        public string RegardingId { get; set; }
        [DataMember]
        public Int16 Status { get; set; } //{0:Active,1:Inactive}
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
        public List<DocumemtType> GetDocumentTypes(DocumemtType docType)
        {
            List<DocumemtType> lstDocumemntTypes= new List<DocumemtType>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (docType.ObjectType.ToString().Trim().Length == 0)
                    {
                        lstDocumemntTypes.Add(new DocumemtType { StatusCode = "204", StatusDescription = "Object Type is required." });
                        return lstDocumemntTypes;
                    }
                    Query = new QueryExpression("hil_tenderattachmentdoctype");
                    Query.ColumnSet = new ColumnSet("hil_name");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_objecttype", ConditionOperator.Equal, docType.ObjectType);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        lstDocumemntTypes.Add(new DocumemtType { StatusCode = "204", StatusDescription = "Object Type does not exist." });
                    }
                    else
                    {
                        foreach (Entity ent in entcoll.Entities)
                        {
                            lstDocumemntTypes.Add(new DocumemtType { ObjectType = docType.ObjectType, Documentype = ent.GetAttributeValue<string>("hil_name"), DocumentypeGuid = ent.Id, StatusCode = "200", StatusDescription = "OK" });
                        }
                    }
                }
                else
                {
                    lstDocumemntTypes.Add(new DocumemtType { StatusCode = "503", StatusDescription = "D365 Service Unavailable" });
                }
            }
            catch (Exception ex)
            {
                lstDocumemntTypes.Add(new DocumemtType { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return lstDocumemntTypes;
        }

        public List<AttachmentResult> GetAttachmentsByRecId(AttachmentManager reqParams)
        {
            List<AttachmentResult> lstDocumemntTypes = new List<AttachmentResult>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (reqParams.RegardingId.ToString().Trim().Length == 0)
                    {
                        lstDocumemntTypes.Add(new AttachmentResult { StatusCode = "204", StatusDescription = "Regarding Id is required." });
                        return lstDocumemntTypes;
                    }
                    Query = new QueryExpression("hil_tenderattachmentmanager");
                    Query.ColumnSet = new ColumnSet(true);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, reqParams.RegardingId);
                    Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, reqParams.Status);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        lstDocumemntTypes.Add(new AttachmentResult { StatusCode = "204", StatusDescription = "Attachment does not exist." });
                    }
                    else
                    {
                        foreach (Entity ent in entcoll.Entities)
                        {
                            lstDocumemntTypes.Add(new AttachmentResult { 
                                CreatedBy= ent.GetAttributeValue<EntityReference>("createdby").Name,
                                DocumentExt= ent.GetAttributeValue<string>("hil_mimetype"),
                                DocumentSize = ent.GetAttributeValue<string>("hil_sizeofdoc"),
                                DocumentType = ent.GetAttributeValue<EntityReference>("hil_documenttype").Name,
                                DocumentTypeGuid = ent.GetAttributeValue<EntityReference>("hil_documenttype").Id,
                                DocumentURL = ent.GetAttributeValue<string>("hil_documenturl"),
                                DocumentVersion = ent.GetAttributeValue<decimal>("hil_version"),
                                LastModifiedOn = ent.GetAttributeValue<DateTime>("modifiedon").AddMinutes(330),
                                RegardingId = reqParams.RegardingId,
                                StatusCode = "200", 
                                StatusDescription = "OK" });
                        }
                    }
                }
                else
                {
                    lstDocumemntTypes.Add(new AttachmentResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" });
                }
            }
            catch (Exception ex)
            {
                lstDocumemntTypes.Add(new AttachmentResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return lstDocumemntTypes;
        }

        public AttachmentManager InsertAttachment(AttachmentResult reqParams)
        {
            AttachmentManager _retObj = null;
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                decimal _version = new decimal(1.00);
                Guid _oldAttachmentId = Guid.Empty;
                if (service != null)
                {
                    if (reqParams.RegardingId.ToString().Trim().Length == 0)
                    {
                        _retObj = new AttachmentManager() { StatusCode = "204", StatusDescription = "Regarding Id is required." };
                        return _retObj;
                    }
                    if (reqParams.DocumentTypeGuid == Guid.Empty)
                    {
                        _retObj = new AttachmentManager() { StatusCode = "204", StatusDescription = "Document Type is required." };
                        return _retObj;
                    }
                    if (reqParams.DocumentURL.ToString().Trim().Length == 0)
                    {
                        _retObj = new AttachmentManager() { StatusCode = "204", StatusDescription = "Document URL is required." };
                        return _retObj;
                    }
                    if (reqParams.DocumentExt.ToString().Trim().Length == 0)
                    {
                        _retObj = new AttachmentManager() { StatusCode = "204", StatusDescription = "Document MIME Type is required." };
                        return _retObj;
                    }
                    if (reqParams.DocumentSize.ToString().Trim().Length == 0)
                    {
                        _retObj = new AttachmentManager() { StatusCode = "204", StatusDescription = "Document Size is required." };
                        return _retObj;
                    }

                    Query = new QueryExpression("hil_tenderattachmentmanager");
                    Query.ColumnSet = new ColumnSet("hil_version");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, reqParams.RegardingId);
                    Query.Criteria.AddCondition("hil_documenttype", ConditionOperator.Equal, reqParams.DocumentTypeGuid);
                    Query.AddOrder("modifiedon", OrderType.Descending);
                    Query.TopCount = 1;
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count > 0)
                    {
                        _version = entcoll.Entities[0].GetAttributeValue<decimal>("hil_version");
                        _version = _version + new decimal(0.01);
                        try
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = entcoll.Entities[0].Id,
                                    LogicalName = "hil_tenderattachmentmanager",
                                },
                                State = new OptionSetValue(1), //Inactive
                                Status = new OptionSetValue(2) //Inactive
                            };
                            service.Execute(setStateRequest);
                        }
                        catch (Exception ex)
                        {
                            _retObj = new AttachmentManager() { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                            return _retObj;
                        }
                    }
                }
                else
                {
                    _retObj = new AttachmentManager() { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                _retObj = new AttachmentManager() { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return _retObj;
        }
    }

    [DataContract]
    public class DocumemtType {
        [DataMember]
        public string ObjectType { get; set; }
        [DataMember]
        public string Documentype { get; set; }
        [DataMember]
        public Guid DocumentypeGuid { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class AttachmentResult {
        [DataMember]
        public string RegardingId { get; set; }
        [DataMember]
        public string DocumentType { get; set; }
        [DataMember]
        public Guid DocumentTypeGuid { get; set; }
        [DataMember]
        public string DocumentURL { get; set; }
        [DataMember]
        public string DocumentExt { get; set; }
        [DataMember]
        public string DocumentSize { get; set; }
        [DataMember]
        public Decimal DocumentVersion { get; set; }
        [DataMember]
        public DateTime LastModifiedOn { get; set; }
        [DataMember]
        public String CreatedBy { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

    }
}
