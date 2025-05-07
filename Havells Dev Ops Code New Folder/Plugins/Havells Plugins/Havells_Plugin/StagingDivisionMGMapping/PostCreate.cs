using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk.Query;


namespace Havells_Plugin.StagingDivisionMGMapping
{
    public class PostCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_stagingdivisonmaterialgroupmapping.EntityLogicalName && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    hil_stagingdivisonmaterialgroupmapping Map = entity.ToEntity<hil_stagingdivisonmaterialgroupmapping>();
                    ResolveLookupsMap(service, Map);

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.StagingDivisionMaterialGroupMapping.PostCreate.Execute" + ex.Message);
            }
        }
        public static void ResolveLookupsMap(IOrganizationService service, hil_stagingdivisonmaterialgroupmapping Map)
        {try
            {
                string Division = string.Empty;
                string MG = string.Empty;
                if ((Map.hil_StagingDivisonUniqueKey != null) && (Map.hil_name != null))
                {
                    Guid fsDivision = GetGuidbyNameProducts(Product.EntityLogicalName, "hil_stagingdivision", Map.hil_StagingDivisonUniqueKey, 2, service);
                    Guid fsMG = GetGuidbyNameProducts(Product.EntityLogicalName, "hil_uniquekey", Map.hil_uniquekey, 3, service);

                    hil_stagingdivisonmaterialgroupmapping upMap = new hil_stagingdivisonmaterialgroupmapping();
                    upMap.hil_stagingdivisonmaterialgroupmappingId = Map.hil_stagingdivisonmaterialgroupmappingId.Value;
                    if (fsDivision != Guid.Empty)
                    {
                        upMap.hil_ProductCategoryDivision = new EntityReference(Product.EntityLogicalName, fsDivision);
                    }
                    if (fsMG != Guid.Empty)
                    {
                        upMap.hil_ProductSubCategoryMG = new EntityReference(Product.EntityLogicalName, fsMG);
                    }
                    service.Update(upMap);

                    if (fsDivision != Guid.Empty && fsMG != Guid.Empty)
                    {
                        SetStateRequest state = new SetStateRequest();
                        state.State = new OptionSetValue(0);
                        state.Status = new OptionSetValue(1);
                        state.EntityMoniker = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Map.hil_stagingdivisonmaterialgroupmappingId.Value);
                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.StagingDivisionMaterialGroupMapping.PostCreate.ResolveLookupsMap" + ex.Message);
            }
        }


        public static Guid GetGuidbyNameProducts(String sEntityName, String sFieldName, String sFieldValue, Int32 iHierarchyLevel, IOrganizationService service, int iStatusCode = 0)
        {
            
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                if (iStatusCode >= 0)
                {
                    // qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                }
                qe.Criteria.AddCondition("hil_deleteflag", ConditionOperator.NotEqual, 1);
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.StagingDivisionMaterialGroupMapping.PostCreate.GetGuidbyNameProducts" + ex.Message);
            }
            return fsResult;
        }

    }
}
