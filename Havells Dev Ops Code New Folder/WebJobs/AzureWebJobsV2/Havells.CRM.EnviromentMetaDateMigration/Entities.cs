using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;

namespace Havells.CRM.EnviromentMetaDateMigration
{
    public class Entities
    {
        static Dictionary<string, string> viewIdsDict = new Dictionary<string, string>();
        public static void EntityMigration(IOrganizationService _SourceService, IOrganizationService _DestinationService)
        {
            WriteLogFile.WriteLog("********************** ENTITY MIGRATEION METHOD **********************");
            string EntityNames = "hil_jobsauth";//"hil_smstemplates;hil_smsconfiguration;invoice;account;contact;incident;msdyn_customerasset;msdyn_workorder;product;systemuser;hil_approvalmatrix;hil_approval;ispl_autonumberconfiguration;hil_autonumber;hil_tenderbankguarantee;hil_branch;hil_businessmapping;hil_casecategory;hil_city;hil_country;hil_district;hil_hsncode;plt_idgenerator;new_integrationstaging;hil_oaheader;hil_oaproduct;hil_orderchecklist;hil_orderchecklistproduct;hil_pincode;hil_region;hil_salesoffice;hil_servicecallrequest;hil_servicecallrequestdetail;hil_state;hil_tender;hil_designteambranchmapping;hil_tenderproduct;msdyn_paymentdetail;msdyn_payment;hil_casecategory;hil_usertype;hil_casecategorymapping;hil_caseassignmentmatrix";
            string[] entities = EntityNames.Split(';');
            WriteLogFile.WriteLog("totla Entity Count " + entities.Length);
            int totalCount = entities.Length;
            int done = 1;
            int error = 0;
            #region createEntity
            foreach (string entityName in entities)
            {
                try
                {
                    WriteLogFile.WriteLog("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@  Enity With Name " + entityName + " started");
                    RetrieveEntityRequest retrieveEntityRequest_Source = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = entityName
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse_Source = (RetrieveEntityResponse)_SourceService.Execute(retrieveEntityRequest_Source);
                    EntityMetadata Entity_Source = retrieveAccountEntityResponse_Source.EntityMetadata;

                    try
                    {
                        RetrieveEntityRequest retrieveEntityRequest_Destination = new RetrieveEntityRequest
                        {
                            EntityFilters = EntityFilters.All,
                            LogicalName = entityName
                        };
                        RetrieveEntityResponse retrieveAccountEntityResponse_Destination = (RetrieveEntityResponse)_DestinationService.Execute(retrieveEntityRequest_Destination);
                        EntityMetadata Entity_Destination = retrieveAccountEntityResponse_Destination.EntityMetadata;

                        updateEntity(_DestinationService, Entity_Source, Entity_Destination);
                        WriteLogFile.WriteLog("Entity Updated " + done + " / " + totalCount);
                        UpsertFieldsExceptLookUp(_DestinationService, Entity_Source);
                        //if (entityName == "hil_casecategory")
                        createNto1RelationShip(_DestinationService, Entity_Source);
                        CreateN2NRelationShip(_DestinationService, Entity_Source);
                        ViewMigration(_SourceService, _DestinationService, Entity_Source.ObjectTypeCode, Entity_Destination.ObjectTypeCode);
                        FormMigration(_SourceService, _DestinationService, Entity_Source.ObjectTypeCode, Entity_Destination.ObjectTypeCode);
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("Could not find an entity with name"))
                        {

                            createEntity(_DestinationService, Entity_Source);
                            UpsertFieldsExceptLookUp(_DestinationService, Entity_Source);
                            createNto1RelationShip(_DestinationService, Entity_Source);
                            CreateN2NRelationShip(_DestinationService, Entity_Source);


                            RetrieveEntityRequest retrieveEntityRequest_Destination = new RetrieveEntityRequest
                            {
                                EntityFilters = EntityFilters.All,
                                LogicalName = entityName
                            };
                            RetrieveEntityResponse retrieveAccountEntityResponse_Destination = (RetrieveEntityResponse)_DestinationService.Execute(retrieveEntityRequest_Destination);
                            EntityMetadata Entity_Destination = retrieveAccountEntityResponse_Destination.EntityMetadata;


                            ViewMigration(_SourceService, _DestinationService, Entity_Source.ObjectTypeCode, Entity_Destination.ObjectTypeCode);
                            FormMigration(_SourceService, _DestinationService, Entity_Source.ObjectTypeCode, Entity_Destination.ObjectTypeCode);
                            WriteLogFile.WriteLog("Entity Created " + done + " / " + totalCount);
                        }
                        else
                            WriteLogFile.WriteLog("Error in Retriving ENtity : " + ex.Message);
                    }
                    done++;
                }
                catch (Exception ex)
                {
                    error++;
                    WriteLogFile.WriteLog("-------------------------Failuer " + error + " in entity creation with name " + entityName + " Error is:-  " + ex.Message);
                }
            }
            #endregion
        }
        public static void createEntity(IOrganizationService _DestinationService, EntityMetadata Entity_Source)
        {
            Label displayName = new Label();
            Label discreption = new Label();
            foreach (AttributeMetadata attr in Entity_Source.Attributes)
            {
                if (attr.LogicalName == Entity_Source.PrimaryNameAttribute)
                {
                    displayName = attr.DisplayName;
                    discreption = attr.Description;
                    goto GenrateMetaDataforCreateEntity;
                }
            }
        GenrateMetaDataforCreateEntity:
            EntityMetadata entityMetadata = new EntityMetadata
            {
                SchemaName = Entity_Source.SchemaName,
                DisplayName = Entity_Source.DisplayName,//new Label("Bank Account", 1033),
                DisplayCollectionName = Entity_Source.DisplayCollectionName,//new Label("Bank Accounts", 1033),
                Description = Entity_Source.Description,//new Label("An entity to store information about customer bank accounts", 1033),
                OwnershipType = Entity_Source.OwnershipType,//OwnershipTypes.UserOwned,
                IsActivity = Entity_Source.IsActivity,
                IsAuditEnabled = Entity_Source.IsAuditEnabled,//
                IsActivityParty = Entity_Source.IsActivityParty,//
                IsBPFEntity = Entity_Source.IsBPFEntity,//
                IsAvailableOffline = Entity_Source.IsAvailableOffline,
                IsBusinessProcessEnabled = Entity_Source.IsBusinessProcessEnabled,
                IsConnectionsEnabled = Entity_Source.IsConnectionsEnabled,
                IsCustomizable = Entity_Source.IsCustomizable,
                IsDocumentManagementEnabled = Entity_Source.IsDocumentManagementEnabled,
                IsDocumentRecommendationsEnabled = Entity_Source.IsDocumentRecommendationsEnabled,
                IsDuplicateDetectionEnabled = Entity_Source.IsDuplicateDetectionEnabled,
                IsEnabledForExternalChannels = Entity_Source.IsEnabledForExternalChannels,
                IsInteractionCentricEnabled = Entity_Source.IsInteractionCentricEnabled,
                IsKnowledgeManagementEnabled = Entity_Source.IsKnowledgeManagementEnabled,
                IsMailMergeEnabled = Entity_Source.IsMailMergeEnabled,
                IsMappable = Entity_Source.IsMappable,
                IsMSTeamsIntegrationEnabled = Entity_Source.IsMSTeamsIntegrationEnabled,
                IsOfflineInMobileClient = Entity_Source.IsOfflineInMobileClient,
                IsOneNoteIntegrationEnabled = Entity_Source.IsOneNoteIntegrationEnabled,
                IsQuickCreateEnabled = Entity_Source.IsQuickCreateEnabled,
                IsReadingPaneEnabled = Entity_Source.IsReadingPaneEnabled,
                ActivityTypeMask = Entity_Source.ActivityTypeMask,
                AutoCreateAccessTeams = Entity_Source.AutoCreateAccessTeams,
                AutoRouteToOwnerQueue = Entity_Source.AutoRouteToOwnerQueue,
                CanChangeHierarchicalRelationship = Entity_Source.CanChangeHierarchicalRelationship,
                CanChangeTrackingBeEnabled = Entity_Source.CanChangeTrackingBeEnabled,
                CanCreateAttributes = Entity_Source.CanCreateAttributes,
                CanCreateCharts = Entity_Source.CanCreateCharts,
                CanCreateForms = Entity_Source.CanCreateForms,
                CanCreateViews = Entity_Source.CanCreateViews,
                CanEnableSyncToExternalSearchIndex = Entity_Source.CanEnableSyncToExternalSearchIndex,
                CanModifyAdditionalSettings = Entity_Source.CanModifyAdditionalSettings,
                ChangeTrackingEnabled = Entity_Source.ChangeTrackingEnabled,
                DataProviderId = Entity_Source.DataProviderId,
                DataSourceId = Entity_Source.DataSourceId,
                DaysSinceRecordLastModified = Entity_Source.DaysSinceRecordLastModified,
                EntityColor = Entity_Source.EntityColor,
                EntityHelpUrl = Entity_Source.EntityHelpUrl,
                EntityHelpUrlEnabled = Entity_Source.EntityHelpUrlEnabled,
                EntitySetName = Entity_Source.EntitySetName,
                ExtensionData = Entity_Source.ExtensionData,
                ExternalCollectionName = Entity_Source.ExternalCollectionName,
                ExternalName = Entity_Source.ExternalName,
                HasActivities = Entity_Source.HasActivities,
                HasChanged = Entity_Source.HasChanged,
                HasEmailAddresses = Entity_Source.HasEmailAddresses,
                HasFeedback = Entity_Source.HasFeedback,
                HasNotes = Entity_Source.HasNotes,
                IconLargeName = Entity_Source.IconLargeName,
                IconMediumName = Entity_Source.IconMediumName,
                IconSmallName = Entity_Source.IconSmallName,
                IconVectorName = Entity_Source.IconVectorName,
                IsReadOnlyInMobileClient = Entity_Source.IsReadOnlyInMobileClient,
                IsRenameable = Entity_Source.IsRenameable,
                IsRetrieveAuditEnabled = Entity_Source.IsRetrieveAuditEnabled,
                IsRetrieveMultipleAuditEnabled = Entity_Source.IsRetrieveMultipleAuditEnabled,
                IsSLAEnabled = Entity_Source.IsSLAEnabled,
                IsSolutionAware = Entity_Source.IsSolutionAware,
                IsValidForQueue = Entity_Source.IsValidForQueue,
                IsVisibleInMobile = Entity_Source.IsVisibleInMobile,
                IsVisibleInMobileClient = Entity_Source.IsVisibleInMobileClient,
                LogicalCollectionName = Entity_Source.LogicalCollectionName,
                LogicalName = Entity_Source.LogicalName,
                //MetadataId = Entity_Source.MetadataId,
                MobileOfflineFilters = Entity_Source.MobileOfflineFilters,
                OwnerId = Entity_Source.OwnerId,
                OwnerIdType = Entity_Source.OwnerIdType,
                OwningBusinessUnit = Entity_Source.OwningBusinessUnit,
                SettingOf = Entity_Source.SettingOf,
                Settings = Entity_Source.Settings,
                SyncToExternalSearchIndex = Entity_Source.SyncToExternalSearchIndex,
                UsesBusinessDataLabelTable = Entity_Source.UsesBusinessDataLabelTable,
                MetadataId = Entity_Source.MetadataId
            };
            if (Entity_Source.IsActivity == true)
            {

                CreateEntityRequest createrequest = new CreateEntityRequest
                {
                    HasActivities = false,
                    HasNotes = true,
                    //Define the entity
                    Entity = entityMetadata,
                    PrimaryAttribute = new StringAttributeMetadata
                    {
                        SchemaName = "Subject",
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.SystemRequired),
                        MaxLength = 100,
                        FormatName = StringFormatName.Text,
                        DisplayName = displayName,
                        Description = discreption
                    }

                };
                _DestinationService.Execute(createrequest);
            }
            else
            {
                CreateEntityRequest createrequest = new CreateEntityRequest
                {
                    //Define the entity
                    HasActivities = (bool)Entity_Source.HasActivities,
                    HasFeedback = (bool)Entity_Source.HasFeedback,
                    HasNotes = (bool)Entity_Source.HasNotes,
                    Entity = entityMetadata,
                    PrimaryAttribute = new StringAttributeMetadata
                    {
                        SchemaName = Entity_Source.PrimaryNameAttribute,
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.SystemRequired),
                        MaxLength = 100,
                        FormatName = StringFormatName.Text,
                        DisplayName = displayName,
                        Description = discreption
                    }
                };
                _DestinationService.Execute(createrequest);
            }
        }
        public static void updateEntity(IOrganizationService _DestinationService, EntityMetadata Entity_Source, EntityMetadata Entity_Destinaiton)
        {
            WriteLogFile.WriteLog("Entity name " + Entity_Source.LogicalName);
            try
            {


                EntityMetadata entityMetadata = new EntityMetadata
                {
                    SchemaName = Entity_Source.SchemaName,
                    DisplayName = Entity_Source.DisplayName,//new Label("Bank Account", 1033),
                    DisplayCollectionName = Entity_Source.DisplayCollectionName,//new Label("Bank Accounts", 1033),
                    Description = Entity_Source.Description,//new Label("An entity to store information about customer bank accounts", 1033),
                    OwnershipType = Entity_Source.OwnershipType,//OwnershipTypes.UserOwned,
                    IsActivity = Entity_Source.IsActivity,
                    IsAuditEnabled = Entity_Source.IsAuditEnabled,//
                    IsActivityParty = Entity_Source.IsActivityParty,//
                    IsBPFEntity = Entity_Source.IsBPFEntity,//
                    IsAvailableOffline = Entity_Source.IsAvailableOffline,
                    IsBusinessProcessEnabled = Entity_Source.IsBusinessProcessEnabled,
                    IsConnectionsEnabled = Entity_Source.IsConnectionsEnabled,
                    IsCustomizable = Entity_Source.IsCustomizable,
                    IsDocumentManagementEnabled = Entity_Source.IsDocumentManagementEnabled,
                    IsDocumentRecommendationsEnabled = Entity_Source.IsDocumentRecommendationsEnabled,
                    IsDuplicateDetectionEnabled = Entity_Source.IsDuplicateDetectionEnabled,
                    IsEnabledForExternalChannels = Entity_Source.IsEnabledForExternalChannels,
                    IsInteractionCentricEnabled = Entity_Source.IsInteractionCentricEnabled,
                    IsKnowledgeManagementEnabled = Entity_Source.IsKnowledgeManagementEnabled,
                    IsMailMergeEnabled = Entity_Source.IsMailMergeEnabled,
                    IsMappable = Entity_Source.IsMappable,
                    IsMSTeamsIntegrationEnabled = Entity_Source.IsMSTeamsIntegrationEnabled,
                    IsOfflineInMobileClient = Entity_Source.IsOfflineInMobileClient,
                    IsOneNoteIntegrationEnabled = Entity_Source.IsOneNoteIntegrationEnabled,
                    IsQuickCreateEnabled = Entity_Source.IsQuickCreateEnabled,
                    IsReadingPaneEnabled = Entity_Source.IsReadingPaneEnabled,
                    ActivityTypeMask = Entity_Source.ActivityTypeMask,
                    AutoCreateAccessTeams = Entity_Source.AutoCreateAccessTeams,
                    AutoRouteToOwnerQueue = Entity_Source.AutoRouteToOwnerQueue,
                    CanChangeHierarchicalRelationship = Entity_Source.CanChangeHierarchicalRelationship,
                    CanChangeTrackingBeEnabled = Entity_Source.CanChangeTrackingBeEnabled,
                    CanCreateAttributes = Entity_Source.CanCreateAttributes,
                    CanCreateCharts = Entity_Source.CanCreateCharts,
                    CanCreateForms = Entity_Source.CanCreateForms,
                    CanCreateViews = Entity_Source.CanCreateViews,
                    CanEnableSyncToExternalSearchIndex = Entity_Source.CanEnableSyncToExternalSearchIndex,
                    CanModifyAdditionalSettings = Entity_Source.CanModifyAdditionalSettings,
                    ChangeTrackingEnabled = Entity_Source.ChangeTrackingEnabled,
                    DataProviderId = Entity_Source.DataProviderId,
                    DataSourceId = Entity_Source.DataSourceId,
                    DaysSinceRecordLastModified = Entity_Source.DaysSinceRecordLastModified,
                    EntityColor = Entity_Source.EntityColor,
                    EntityHelpUrl = Entity_Source.EntityHelpUrl,
                    EntityHelpUrlEnabled = Entity_Source.EntityHelpUrlEnabled,
                    EntitySetName = Entity_Source.EntitySetName,
                    ExtensionData = Entity_Source.ExtensionData,
                    ExternalCollectionName = Entity_Source.ExternalCollectionName,
                    ExternalName = Entity_Source.ExternalName,
                    HasActivities = Entity_Source.HasActivities,
                    HasChanged = Entity_Source.HasChanged,
                    HasEmailAddresses = Entity_Source.HasEmailAddresses,
                    HasFeedback = Entity_Source.HasFeedback,
                    HasNotes = Entity_Source.HasNotes,
                    IconLargeName = Entity_Source.IconLargeName,
                    IconMediumName = Entity_Source.IconMediumName,
                    IconSmallName = Entity_Source.IconSmallName,
                    IconVectorName = Entity_Source.IconVectorName,
                    IsReadOnlyInMobileClient = Entity_Source.IsReadOnlyInMobileClient,
                    IsRenameable = Entity_Source.IsRenameable,
                    IsRetrieveAuditEnabled = Entity_Source.IsRetrieveAuditEnabled,
                    IsRetrieveMultipleAuditEnabled = Entity_Source.IsRetrieveMultipleAuditEnabled,
                    IsSLAEnabled = Entity_Source.IsSLAEnabled,
                    IsSolutionAware = Entity_Source.IsSolutionAware,
                    IsValidForQueue = Entity_Source.IsValidForQueue,
                    IsVisibleInMobile = Entity_Source.IsVisibleInMobile,
                    IsVisibleInMobileClient = Entity_Source.IsVisibleInMobileClient,
                    LogicalCollectionName = Entity_Source.LogicalCollectionName,
                    LogicalName = Entity_Source.LogicalName,
                    MetadataId = Entity_Source.MetadataId,
                    MobileOfflineFilters = Entity_Source.MobileOfflineFilters,
                    OwnerId = Entity_Source.OwnerId,
                    OwnerIdType = Entity_Source.OwnerIdType,
                    OwningBusinessUnit = Entity_Source.OwningBusinessUnit,
                    SettingOf = Entity_Source.SettingOf,
                    Settings = Entity_Source.Settings,
                    SyncToExternalSearchIndex = Entity_Source.SyncToExternalSearchIndex,
                    UsesBusinessDataLabelTable = Entity_Source.UsesBusinessDataLabelTable
                };
                entityMetadata.MetadataId = Entity_Destinaiton.MetadataId;
                UpdateEntityRequest updatereq = new UpdateEntityRequest();
                updatereq.HasActivities = (bool)Entity_Source.HasActivities;
                updatereq.HasFeedback = (bool)Entity_Source.HasFeedback;
                updatereq.HasNotes = (bool)Entity_Source.HasNotes;
                updatereq.Entity = entityMetadata;

                UpdateEntityResponse updateresp = (UpdateEntityResponse)_DestinationService.Execute(updatereq);
                WriteLogFile.WriteLog("Entity Updated ");

            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog("------------------------- ENTITY UPDATION FAILDE WITH NAME " + Entity_Source.LogicalName + " with ERROR " + ex.Message);
            }

        }
        public static void UpsertFieldsExceptLookUp(IOrganizationService _DestinationService, EntityMetadata Entity_Source)
        {
            if (!Entity_Source.LogicalName.Contains("archived"))
            {
                WriteLogFile.WriteLog("********************** FIELD CREATION FOR ENTITY " + Entity_Source.LogicalName.ToUpper() + " IS STARTED **********************");
                WriteLogFile.WriteLog("totla Attribute Count " + Entity_Source.Attributes.Length);
                int totalCount = Entity_Source.Attributes.Length;
                int done = 1;
                int error = 0;
                RetrieveEntityRequest retrieveEntityRequest_Destinaiton = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = Entity_Source.LogicalName
                };
                RetrieveEntityResponse retrieveAccountEntityResponse_Source = (RetrieveEntityResponse)_DestinationService.Execute(retrieveEntityRequest_Destinaiton);
                EntityMetadata Entity_Destinaiton = retrieveAccountEntityResponse_Source.EntityMetadata;

                AttributeMetadata[] Attribute_Destnation_Array = Entity_Destinaiton.Attributes;

                foreach (AttributeMetadata sourceAttribue in Entity_Source.Attributes)
                {
                    try
                    {

                        if (!sourceAttribue.LogicalName.Contains("_base") && sourceAttribue.IsValidForForm == true
                                && sourceAttribue.LogicalName != Entity_Source.PrimaryNameAttribute && sourceAttribue.AttributeType != AttributeTypeCode.Lookup
                                && sourceAttribue.AttributeType != AttributeTypeCode.Customer && sourceAttribue.AttributeType != AttributeTypeCode.Uniqueidentifier)
                        {
                            //var destnationAttribute_found = from Attribute_Destnation in Attribute_Destnation_Array
                            //                                where Attribute_Destnation.SchemaName.ToLower() == sourceAttribue.SchemaName.ToLower()
                            //                                select Attribute_Destnation;

                            AttributeMetadata destnationAttribute = Array.Find(Attribute_Destnation_Array, ele => ele.SchemaName.ToLower() == sourceAttribue.SchemaName.ToLower());

                            AttributeMetadata destnationAttributeNew = sourceAttribue;
                            if (sourceAttribue.SchemaName.Contains("_"))
                                if (sourceAttribue.AttributeTypeName.Value == "MultiSelectPicklistType")
                                {
                                    OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)sourceAttribue).OptionSet;
                                    if (optionSetMetadata.IsGlobal == true)
                                        optionSetMetadata.Options.Clear();
                                    ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)destnationAttributeNew).OptionSet = optionSetMetadata;
                                }
                                else if (sourceAttribue.AttributeType == AttributeTypeCode.Picklist)
                                {
                                    OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)sourceAttribue).OptionSet;
                                    if (optionSetMetadata.IsGlobal == true)
                                        optionSetMetadata.Options.Clear();
                                    ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)destnationAttributeNew).OptionSet = optionSetMetadata;

                                }

                            if (destnationAttribute != null)
                            {
                                destnationAttributeNew.MetadataId = destnationAttribute.MetadataId;
                                UpdateAttributeRequest createAttributeRequest = new UpdateAttributeRequest
                                {
                                    EntityName = Entity_Source.LogicalName,
                                    Attribute = destnationAttributeNew
                                };
                                try
                                {
                                    _DestinationService.Execute(createAttributeRequest);
                                    WriteLogFile.WriteLog("attribute Created " + done + " / " + totalCount + " with name " + destnationAttributeNew.LogicalName);
                                    done++;
                                }
                                catch (Exception ex)
                                {
                                    error++;
                                    WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with name " + sourceAttribue.LogicalName +
                                        " Error is:-  " + ex.Message);
                                }
                            }
                            else
                            {
                                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                                {
                                    EntityName = Entity_Source.LogicalName,
                                    Attribute = sourceAttribue
                                };
                                try
                                {
                                    _DestinationService.Execute(createAttributeRequest);
                                    WriteLogFile.WriteLog("attribute Created " + done + " / " + totalCount + " with name " + sourceAttribue.LogicalName);
                                    done++;
                                }
                                catch (Exception ex)
                                {
                                    error++;
                                    WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with name " + sourceAttribue.LogicalName + " Error is:-  " + ex.Message);
                                }
                            }
                        }
                        if (sourceAttribue.SchemaName.ToLower() == "statuscode")
                        {
                            if (sourceAttribue.AttributeType == AttributeTypeCode.Status)
                            {

                                RetrieveEntityRequest retrieveEntityRequestSer = new RetrieveEntityRequest
                                {
                                    EntityFilters = EntityFilters.All,
                                    LogicalName = Entity_Source.LogicalName
                                };
                                RetrieveEntityResponse retrieveAccountEntityResponseSer = (RetrieveEntityResponse)_DestinationService.Execute(retrieveEntityRequestSer);
                                EntityMetadata sourceAttribueSer = retrieveAccountEntityResponseSer.EntityMetadata;

                                var aa = sourceAttribueSer.Attributes;

                                StatusAttributeMetadata lookField = (StatusAttributeMetadata)Array.Find(aa, ele => ele.SchemaName.ToLower() == sourceAttribue.SchemaName.ToLower());

                                OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)sourceAttribue).OptionSet;

                                sourceAttribue.MetadataId = lookField.MetadataId;

                                lookField.OptionSet = optionSetMetadata;


                                UpdateAttributeRequest createAttributeRequest = new UpdateAttributeRequest
                                {
                                    EntityName = Entity_Source.LogicalName,
                                    Attribute = lookField
                                };
                                try
                                {
                                    _DestinationService.Execute(createAttributeRequest);
                                    WriteLogFile.WriteLog("attribute Created " + done + " / " + totalCount + " with name " + sourceAttribue.LogicalName);
                                    done++;
                                }
                                catch (Exception ex)
                                {
                                    error++;
                                    WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with name " + sourceAttribue.LogicalName + " Error is:-  " + ex.Message);
                                }
                            }
                        }
                        else if (sourceAttribue.SchemaName.ToLower() == "statecode")
                        {
                            if (sourceAttribue.AttributeType == AttributeTypeCode.State)
                            {

                                RetrieveEntityRequest retrieveEntityRequestSer = new RetrieveEntityRequest
                                {
                                    EntityFilters = EntityFilters.All,
                                    LogicalName = Entity_Source.LogicalName
                                };
                                RetrieveEntityResponse retrieveAccountEntityResponseSer = (RetrieveEntityResponse)_DestinationService.Execute(retrieveEntityRequestSer);
                                EntityMetadata sourceAttribueSer = retrieveAccountEntityResponseSer.EntityMetadata;

                                var aa = sourceAttribueSer.Attributes;

                                StateAttributeMetadata lookField = (StateAttributeMetadata)Array.Find(aa, ele => ele.SchemaName.ToLower() == sourceAttribue.SchemaName.ToLower());

                                OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)sourceAttribue).OptionSet;

                                sourceAttribue.MetadataId = lookField.MetadataId;

                                lookField.OptionSet = optionSetMetadata;


                                UpdateAttributeRequest createAttributeRequest = new UpdateAttributeRequest
                                {
                                    EntityName = Entity_Source.LogicalName,
                                    Attribute = lookField
                                };
                                try
                                {
                                    _DestinationService.Execute(createAttributeRequest);
                                    WriteLogFile.WriteLog("attribute Created " + done + " / " + totalCount + " with name " + sourceAttribue.LogicalName);
                                    done++;
                                }
                                catch (Exception ex)
                                {
                                    error++;
                                    WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with name " + sourceAttribue.LogicalName + " Error is:-  " + ex.Message);
                                }
                            }
                        }
                        else
                        {
                            WriteLogFile.WriteLog("attribute Skiped " + done + " / " + totalCount + " with name " + sourceAttribue.LogicalName);
                            done++;
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with name " + sourceAttribue.LogicalName +
                                   " Error is:-  " + ex.Message);
                    }
                }
                WriteLogFile.WriteLog("********************** FIELD CREATION FOR ENTITY " + Entity_Source.LogicalName.ToUpper() + " IS ENDED **********************");

            }
        }
        public static void createNto1RelationShip(IOrganizationService _DestinationService, EntityMetadata Entity_Source)
        {
            WriteLogFile.WriteLog("********************** N to 1 RELATIONSHIP CREATION FOR ENTITY " + Entity_Source.LogicalName.ToUpper() + " IS STARTED **********************");
            WriteLogFile.WriteLog("totla Attribute Count " + Entity_Source.ManyToOneRelationships.Length);
            int totalCount = Entity_Source.ManyToOneRelationships.Length;
            int done = 1;
            int error = 0;
            foreach (OneToManyRelationshipMetadata relationShip_source in Entity_Source.ManyToOneRelationships)
            {
                try
                {
                    if (relationShip_source.ReferencingAttribute.Contains("_"))
                    {
                        AttributeMetadata[] a = Entity_Source.Attributes;
                        AttributeMetadata lookField = Array.Find(a, ele => ele.SchemaName.ToLower() == relationShip_source.ReferencingAttribute.ToLower());
                        relationShip_source.ReferencingAttribute = null;
                        if (lookField.AttributeType == AttributeTypeCode.Customer)
                            continue;
                        CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
                        {
                            OneToManyRelationship = relationShip_source,
                            Lookup = (LookupAttributeMetadata)lookField
                        };
                        CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)_DestinationService.Execute(createOneToManyRelationshipRequest);
                        WriteLogFile.WriteLog("attribute Created " + done + " / " + totalCount + " with Lookup name " + lookField.LogicalName);
                        done++;
                    }
                    else
                        WriteLogFile.WriteLog("RelatoinShip with ReferencingAttribute name " + relationShip_source.ReferencingAttribute + " is skiped due to Default field");

                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("not unique within an entity"))
                    {
                        error++;
                        WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with ReferencingAttribute name " + relationShip_source.ReferencingAttribute + " Error is:-  " + ex.Message);
                    }
                }

            }
            WriteLogFile.WriteLog("********************** N to 1 RELATIONSHIP CREATION FOR ENTITY " + Entity_Source.LogicalName.ToUpper() + " IS ENDED **********************");

        }
        public static void CreateN2NRelationShip(IOrganizationService _DestinationService, EntityMetadata Entity_Source)
        {
            WriteLogFile.WriteLog("********************** N to N RELATIONSHIP CREATION FOR ENTITY " + Entity_Source.LogicalName.ToUpper() + " IS STARTED **********************");
            WriteLogFile.WriteLog("totla Attribute Count " + Entity_Source.ManyToManyRelationships.Length);
            int totalCount = Entity_Source.ManyToManyRelationships.Length;
            int done = 1;
            int error = 0;
            foreach (object relationShip in Entity_Source.ManyToManyRelationships)
            {
                ManyToManyRelationshipMetadata relation = (ManyToManyRelationshipMetadata)relationShip;
                try
                {
                    CreateManyToManyRequest createManyToManyRelationshipRequest = new CreateManyToManyRequest
                    {
                        IntersectEntitySchemaName = relation.IntersectEntityName,
                        ManyToManyRelationship = relation
                    };

                    CreateManyToManyResponse createManytoManyRelationshipResponse = (CreateManyToManyResponse)_DestinationService.Execute(createManyToManyRelationshipRequest);
                    WriteLogFile.WriteLog("RelationShip is Created " + done + " / " + totalCount + " with IntersectEntityName name " + relation.IntersectEntityName);
                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("s not unique within an entity"))
                    {
                        error++;
                        WriteLogFile.WriteLog("-------------------------Failuer " + error + " in attribute creation with IntersectEntityName name " + relation.IntersectEntityName + " Error is:-  " + ex.Message);
                    }
                }

            }
            WriteLogFile.WriteLog("********************** N to N RELATIONSHIP CREATION FOR ENTITY " + Entity_Source.LogicalName.ToUpper() + " IS ENDED **********************");

        }
        public static void FormMigration(IOrganizationService _SourceService, IOrganizationService _DestinationService, int? sourceEntityObjectType, int? destinationEntityObjectType)
        {
            try
            {
                var query = new QueryExpression("systemform");
                query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, sourceEntityObjectType);
                query.ColumnSet = new ColumnSet(true);
                EntityCollection sourceSystemFormCollection = _SourceService.RetrieveMultiple(query);
                //Console.Clear();
                foreach (Entity sourceSystemForm in sourceSystemFormCollection.Entities)
                {
                    try
                    {
                        var query1 = new QueryExpression("systemform");
                        query1.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, destinationEntityObjectType);
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, sourceSystemForm.GetAttributeValue<string>("name"));
                        query1.Criteria.AddCondition("type", ConditionOperator.Equal, sourceSystemForm.GetAttributeValue<OptionSetValue>("type").Value);
                        //entity.FormattedValues["type"]
                        query1.ColumnSet = new ColumnSet(true);
                        EntityCollection destinationSystemFormCollection = _DestinationService.RetrieveMultiple(query1);
                        WriteLogFile.WriteLog(sourceSystemForm.GetAttributeValue<string>("name") + " found " + destinationSystemFormCollection.Entities.Count);
                        if (destinationSystemFormCollection.Entities.Count > 0)
                        {
                            string formXML = sourceSystemForm.GetAttributeValue<string>("formxml");
                            string formJSON = sourceSystemForm.GetAttributeValue<string>("formjson");
                            viewIdsDict = new Dictionary<string, string>();
                            ChaneFormXml(formXML, _SourceService, _DestinationService);

                            foreach (string key in viewIdsDict.Keys)
                            {
                                // WriteLogFile.WriteLog(key + " : " + viewIdsDict[key]);
                                formXML = formXML.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                                formJSON = formJSON.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                            }
                            destinationSystemFormCollection[0]["description"] = sourceSystemForm.GetAttributeValue<string>("description");
                            destinationSystemFormCollection[0]["formactivationstate"] = sourceSystemForm.GetAttributeValue<OptionSetValue>("formactivationstate");
                            destinationSystemFormCollection[0]["formjson"] = formJSON;// entity.GetAttributeValue<string>("formjson");
                            destinationSystemFormCollection[0]["formpresentation"] = sourceSystemForm.GetAttributeValue<OptionSetValue>("formpresentation");
                            destinationSystemFormCollection[0]["formxml"] = formXML; //entity.GetAttributeValue<string>("formxml");
                            destinationSystemFormCollection[0]["isdefault"] = sourceSystemForm.GetAttributeValue<bool>("isdefault");
                            destinationSystemFormCollection[0]["isdesktopenabled"] = sourceSystemForm.GetAttributeValue<bool>("isdesktopenabled");
                            destinationSystemFormCollection[0]["istabletenabled"] = sourceSystemForm.GetAttributeValue<bool>("istabletenabled");
                            destinationSystemFormCollection[0]["type"] = sourceSystemForm.GetAttributeValue<OptionSetValue>("type");
                            destinationSystemFormCollection[0]["version"] = sourceSystemForm.GetAttributeValue<int>("version");
                            destinationSystemFormCollection[0]["name"] = sourceSystemForm.GetAttributeValue<string>("name");
                            _DestinationService.Update(destinationSystemFormCollection[0]);
                        }
                        else
                        {
                            _DestinationService.Create(sourceSystemForm);
                        }
                    }
                    catch (Exception ex)
                    {
                         WriteLogFile.WriteLog("============Error while processing Form " + sourceSystemForm.GetAttributeValue<string>("name") + "  ||  Erorr " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog("Error " + ex.Message);
            }
        }
        public static void ViewMigration(IOrganizationService _SourceService, IOrganizationService _DestinationService, int? sourceEntityObjectType, int? destinationEntityObjectType)
        {
            try
            {
                var query = new QueryExpression("savedquery");
                query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, sourceEntityObjectType);
                query.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
                RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)_SourceService.Execute(retrieveSavedQueriesRequest);
                DataCollection<Entity> SourceSavedQueriesColl = retrieveSavedQueriesResponse.EntityCollection.Entities;
                foreach (Entity SourceSavedQuery in SourceSavedQueriesColl)
                {
                    try
                    {

                        var queryDemo = new QueryExpression("savedquery");
                        queryDemo.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, destinationEntityObjectType);
                        queryDemo.Criteria.AddCondition("querytype", ConditionOperator.Equal, SourceSavedQuery.GetAttributeValue<int>("querytype"));
                        queryDemo.Criteria.AddCondition("name", ConditionOperator.Equal, SourceSavedQuery.GetAttributeValue<string>("name"));
                        queryDemo.ColumnSet = new ColumnSet(true);
                        RetrieveMultipleRequest retrieveSavedQueriesRequestDemo = new RetrieveMultipleRequest { Query = queryDemo };
                        RetrieveMultipleResponse retrieveSavedQueriesResponseDemo = (RetrieveMultipleResponse)_DestinationService.Execute(retrieveSavedQueriesRequestDemo);
                        DataCollection<Entity> DestinationSavedQueryColl = retrieveSavedQueriesResponseDemo.EntityCollection.Entities;
                        if (DestinationSavedQueryColl.Count == 0)
                        {
                            _DestinationService.Create(SourceSavedQuery);
                            WriteLogFile.WriteLog("more than 1 rec");
                        }
                        else
                        {
                            foreach (Entity DestinationSavedQuery in DestinationSavedQueryColl)
                            {
                                try
                                {
                                    Entity entity = new Entity(DestinationSavedQuery.LogicalName, DestinationSavedQuery.Id);
                                    entity["columnsetxml"] = SourceSavedQuery.GetAttributeValue<string>("columnsetxml");
                                    entity["fetchxml"] = SourceSavedQuery.GetAttributeValue<string>("fetchxml");
                                    entity["layoutjson"] = SourceSavedQuery.GetAttributeValue<string>("layoutjson");
                                    entity["layoutxml"] = SourceSavedQuery.GetAttributeValue<string>("layoutxml");
                                    entity["name"] = SourceSavedQuery.GetAttributeValue<string>("name");
                                    entity["description"] = SourceSavedQuery.GetAttributeValue<string>("description");
                                    _DestinationService.Update(entity);
                                    WriteLogFile.WriteLog(SourceSavedQuery.GetAttributeValue<string>("name") + " found " + DestinationSavedQueryColl.Count);
                                }
                                catch (Exception ex)
                                {
                                    WriteLogFile.WriteLog("----- Error in Updating View with Name  " + SourceSavedQuery.GetAttributeValue<string>("name") + " || Error is " + ex.Message);
                                }

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLogFile.WriteLog("----- Error in Creation View with Name  " + SourceSavedQuery.GetAttributeValue<string>("name") + " || Error is " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog("Error " + ex.Message);
            }
        }
        public static void ChaneFormXml(string xml, IOrganizationService service, IOrganizationService serviceSer)
        {
            try
            {
                xml = xml.Replace("<ViewIds>", "|");
                xml = xml.Replace("</ViewIds>", "^");
                xml = xml.Replace("<AvailableViewIds>", "|");
                xml = xml.Replace("</AvailableViewIds>", "^");
                xml = xml.Replace("<DefaultViewId>", "|");
                xml = xml.Replace("</DefaultViewId>", "^");
                string[] allViews = xml.Split('|');
                foreach (string a in allViews)
                {
                    string[] views = a.Split('^');
                    if (views.Length == 1)
                        continue;
                    string[] vewsIds = views[0].Split(',');
                    foreach (string viewid in vewsIds)
                    {
                        var dictval = from x in viewIdsDict
                                      where x.Key.Contains(viewid)
                                      select x;
                        if (dictval.ToList().Count == 0)
                        {
                            string newViewId = getDemoViewId(viewid, service, serviceSer);
                            newViewId = newViewId == null ? "00000000-0000-0000-0000-000000000000" : newViewId;
                            viewIdsDict.Add(viewid, "{" + newViewId + "}");
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog(ex.Message);
            }
        }
        public static string getDemoViewId(string ViewIdPrd, IOrganizationService service, IOrganizationService serviceSer)
        {

            string newViewId = null;
            try
            {
                var query = new QueryExpression("savedquery");
                query.Criteria.AddCondition("savedqueryid", ConditionOperator.Equal, ViewIdPrd);
                query.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
                RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
                DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;

                var query1 = new QueryExpression("savedquery");
                query1.Criteria.AddCondition("name", ConditionOperator.Equal, savedQueries[0].GetAttributeValue<string>("name"));
                query1.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;

                foreach (Entity ent in savedQueries1)
                {
                    if (ent.GetAttributeValue<string>("returnedtypecode") == savedQueries[0].GetAttributeValue<string>("returnedtypecode"))
                    {
                        newViewId = ent.Id.ToString();
                    }
                }
                if (savedQueries1.Count == 0)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    WriteLogFile.WriteLog(nana);
                }
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("The entity with a name ="))
                    WriteLogFile.WriteLog(ex.Message);
            }
            if (newViewId == null)
                WriteLogFile.WriteLog("dd");
            return newViewId;
        }
    }
}
