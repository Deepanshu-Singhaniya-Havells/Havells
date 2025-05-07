using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.SystemUsers
{
    public class PostUpdate_BusinessUnit : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            QueryExpression query = null;
            EntityCollection entColl = null;
            QueryExpression query1 = null;
            EntityCollection entColl1 = null;
            Guid businessUnitGuid = Guid.Empty;
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "systemuser" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    query = new QueryExpression(SystemUser.EntityLogicalName);
                    query.ColumnSet = new ColumnSet("businessunitid");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, entity.Id));
                    entColl = service.RetrieveMultiple(query);
                    if (entColl.Entities.Count > 0)
                    {
                        businessUnitGuid = entColl.Entities[0].GetAttributeValue<EntityReference>("businessunitid").Id;
                        #region Flush User Security Role Extension data
                        query1 = new QueryExpression("hil_usersecurityroleextension");
                        query1.ColumnSet = new ColumnSet("hil_usersecurityroleextensionid");
                        query1.Criteria = new FilterExpression(LogicalOperator.And);
                        query1.Criteria.AddCondition(new ConditionExpression("hil_user", ConditionOperator.Equal, entity.Id));
                        query1.Criteria.AddCondition(new ConditionExpression("hil_businessunit", ConditionOperator.Equal, businessUnitGuid));
                        entColl1 = service.RetrieveMultiple(query1);
                        if (entColl1.Entities.Count > 0)
                        {
                            foreach (Entity ent in entColl1.Entities)
                            {
                                //service.Delete("hil_usersecurityroleextension", ent.Id);
                            }
                        }
                        #endregion
                        #region Resync Security Role Extension Data
                        query = new QueryExpression("role");
                        query.ColumnSet = new ColumnSet("roleid", "name");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitGuid));
                        entColl = service.RetrieveMultiple(query);
                        if (entColl.Entities.Count > 0)
                        {
                            Entity entSRExt;
                            foreach (Entity ent in entColl.Entities)
                            {
                                query1 = new QueryExpression("hil_securityroleextension");
                                query1.ColumnSet = new ColumnSet("hil_securityroleextensionid");
                                query1.Criteria = new FilterExpression(LogicalOperator.And);
                                query1.Criteria.AddCondition(new ConditionExpression("hil_securityrole", ConditionOperator.Equal, ent.Id));
                                query1.Criteria.AddCondition(new ConditionExpression("hil_businessunit", ConditionOperator.Equal, businessUnitGuid));
                                entColl1 = service.RetrieveMultiple(query1);
                                if (entColl1.Entities.Count == 0)
                                {
                                    entSRExt = new Entity("hil_securityroleextension");
                                    entSRExt["hil_securityrole"] = new EntityReference("role", ent.Id);
                                    entSRExt["hil_businessunit"] = new EntityReference("businessunit", businessUnitGuid);
                                    entSRExt["hil_name"] = ent.GetAttributeValue<string>("name");
                                    service.Create(entSRExt);
                                }
                            }
                        }
                        #endregion
                    }
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SystemUsers.PostUpdate_BusinessUnit " + e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SystemUsers.PostUpdate_BusinessUnit" + e.Message);
            }
            #endregion
        }
    }
}
