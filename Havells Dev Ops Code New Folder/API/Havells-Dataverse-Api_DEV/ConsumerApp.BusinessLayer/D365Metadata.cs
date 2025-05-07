using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.Crm.Sdk.Messages;

namespace ConsumerApp.BusinessLayer
{
    public class D365Metadata
    {
        public GlobalSearchResponse GetGlobalSearch(GlobalSearchRequest req)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            GlobalSearchResponse globalSearchResponse = new GlobalSearchResponse();
            try
            {
                List<GlobalSearchEntity> globalSearchEntities = new List<GlobalSearchEntity>();
                List<GlobalSearchEntity> globalSearchEntitiesTemp = new List<GlobalSearchEntity>();
                globalSearchEntities.Add(new GlobalSearchEntity { EntityDisplayName = "None", EntityID = Guid.Empty, EntityLogicalName = "None" });
                QuickFindConfigurationCollection quickFindConfigurations = new QuickFindConfigurationCollection();
                var request = new OrganizationRequest("RetrieveEntityGroupConfiguration");
                request.Parameters["EntityGroupName"] = "Mobile Client Search";
                var response = service.Execute(request);
                quickFindConfigurations = (QuickFindConfigurationCollection)response.Results["EntityGroupConfiguration"];
                foreach (QuickFindConfiguration quickFind in quickFindConfigurations)
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.Entity,
                        LogicalName = quickFind.EntityName
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
                    GlobalSearchEntity globalSearchEntity = new GlobalSearchEntity();
                    globalSearchEntity.EntityDisplayName = StateEntity.DisplayName.LocalizedLabels[0].Label.ToString();
                    globalSearchEntity.EntityLogicalName = StateEntity.LogicalName;
                    globalSearchEntity.EntityID = StateEntity.MetadataId;
                    globalSearchEntitiesTemp.Add(globalSearchEntity);
                }
                RetrieveAppComponentsRequest retrieveAppComponents = new RetrieveAppComponentsRequest();
                retrieveAppComponents.AppModuleId = new Guid(req.AppId);
                var retrieveAppComponentsResponse = service.Execute(retrieveAppComponents);
                EntityCollection entCol = (EntityCollection)retrieveAppComponentsResponse.Results.Values.FirstOrDefault();
                foreach (GlobalSearchEntity quickFind1 in globalSearchEntitiesTemp)
                {
                    var a = entCol.Entities.Where(X => X.GetAttributeValue<Guid>("objectid") == quickFind1.EntityID).ToList();
                    if (a.Count == 1)
                    {
                        globalSearchEntities.Add(quickFind1);
                        Console.WriteLine(quickFind1.EntityDisplayName);
                    }
                }
                globalSearchResponse.GlobalSearchEntities = globalSearchEntities;
                globalSearchResponse.Status = "Sucess";
                globalSearchResponse.ErrorMessage = null;
            }
            catch (Exception ex)
            {
                globalSearchResponse.Status = "Error";
                globalSearchResponse.ErrorMessage = ex.Message;
            }
            return globalSearchResponse;
        }
    }

    [DataContract]
    public class GlobalSearchEntity
    {
        [DataMember]
        public string EntityDisplayName { get; set; }
        [DataMember] public string EntityLogicalName { get; set; }
        [DataMember] public Guid? EntityID { get; set; }
    }

    [DataContract]
    public class GlobalSearchResponse
    {
        [DataMember]
        public List<GlobalSearchEntity> GlobalSearchEntities { get; set; }
        [DataMember] 
        public string Status { get; set; }
        [DataMember] 
        public string ErrorMessage { get; set; }
    }
    [DataContract]
    public class GlobalSearchRequest
    {
        [DataMember]
        public string AppId { get; set; }
    }
}
