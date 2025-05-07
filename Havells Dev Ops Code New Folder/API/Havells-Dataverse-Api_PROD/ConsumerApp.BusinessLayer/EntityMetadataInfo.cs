using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class EntityMetadataInfo
    {
        [DataMember]
        public string LogicalName { get; set; }

        public List<AttributeMetadataInfo> AttributeMetadata(EntityMetadataInfo entity)
        {
            List<AttributeMetadataInfo> attributesList = new List<AttributeMetadataInfo>();
            AttributeMetadataInfo attributesMetadata;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.Attributes,
                        LogicalName = entity.LogicalName
                    };
                    RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata entityMetadata = retrieveEntityResponse.EntityMetadata;
                    foreach (object attribute in entityMetadata.Attributes)
                    {
                        AttributeMetadata a = (AttributeMetadata)attribute;
                        try
                        {
                            attributesMetadata = new AttributeMetadataInfo()
                            {
                                LogicalName = a.LogicalName,
                                DisplayName = a.DisplayName.UserLocalizedLabel == null ? a.SchemaName : a.DisplayName.UserLocalizedLabel.Label
                            };
                            attributesList.Add(attributesMetadata);
                        }
                        catch {}
                    }
                    return attributesList;
                }
            }
            catch {}
            return attributesList;
        }
    }

    [DataContract]
    public class AttributeMetadataInfo
    {
        [DataMember]
        public string LogicalName { get; set; }

        [DataMember]
        public string DisplayName { get; set; }
    }
}
