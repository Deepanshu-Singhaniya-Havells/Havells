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
using System.Collections.Generic;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class Tender
    {
        public List<DocumentTypes> GetDocumentTypes()
        {
            List<DocumentTypes> _retObj = new List<DocumentTypes>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                if (service != null)
                {
                    QueryExpression qryExp = new QueryExpression("hil_enquirydocumenttype");
                    qryExp.ColumnSet = new ColumnSet("hil_name");
                    qryExp.AddOrder("hil_name", OrderType.Ascending);
                    qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    EntityCollection entCol = service.RetrieveMultiple(qryExp);
                    if (entCol.Entities.Count > 0) {
                        foreach (Entity ent in entCol.Entities) {
                            _retObj.Add(new DocumentTypes() { DocumentGuid = ent.Id, DocumentType = ent.GetAttributeValue<string>("hil_name"), StatusCode = "200", StatusDescription = "OK" });
                        }
                    }
                }
                else
                {
                    _retObj.Add(new DocumentTypes { StatusCode = "503", StatusDescription = "D365 Service Unavailable" });
                }
            }
            catch (Exception ex)
            {
                _retObj.Add(new DocumentTypes { StatusCode = "503", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return _retObj;
        }
    }

    [DataContract]
    public class DocumentTypes
    {
        [DataMember]
        public Guid DocumentGuid { get; set; }
        [DataMember]
        public string DocumentType { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
