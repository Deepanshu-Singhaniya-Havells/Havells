using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.UserSecurityRoleExtension
{
    public class PostCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            QueryExpression query = null;
            EntityCollection entColl = null;

            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_usersecurityroleextension" && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    LinkEntity le = new LinkEntity("hil_usersecurityroleextension", "hil_securityroleextension", "hil_securityroleextensionid", "hil_securityroleextension", JoinOperator.Inner);

                    query = new QueryExpression("hil_usersecurityroleextension");
                    query.ColumnSet = new ColumnSet("hil_user");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_usersecurityroleextensionid", ConditionOperator.Equal, entity.Id));
                    entColl = service.RetrieveMultiple(query);
                    if (entColl.Entities.Count > 0)
                    {
                        
                    }
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.UserSecurityRoleExtension.PostCreate " + e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.UserSecurityRoleExtension.PostCreate " + e.Message);
            }
            #endregion
        }
    }
}
