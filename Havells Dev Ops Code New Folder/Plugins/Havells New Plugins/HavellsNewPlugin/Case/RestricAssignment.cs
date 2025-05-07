using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    public class RestricAssignment : IPlugin
    {
        private IOrganizationService service;
        private ITracingService tracingService = null;
        private IPluginExecutionContext context;
        private readonly Guid CRMAdmin = new Guid("5190416c-0782-e911-a959-000d3af06a98");

        private readonly List<string> allowedRoles = new List<string>() { "System Administrator", "Grievance MDM" };
        // This method is used to restrict the users to assign the activity to another user. 
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.MessageName == "Update" && context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "hil_grievancehandlingactivity")
                {
                    Entity targetEntity = (Entity)context.InputParameters["Target"];
                    if (targetEntity.Contains("ownerid"))
                    {
                        Guid userId = context.InitiatingUserId;
                        if (userId != CRMAdmin)
                        {
                            List<string> userRoles = GetUserRoles(service, userId);
                            if (!userRoles.Any(role => allowedRoles.Contains(role)))
                            {
                                throw new InvalidPluginExecutionException("You are not authorized to re-assign the activity");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"Method:{this.GetType().Name} Error:{ex.Message}");
            }

        }
        public List<string> GetUserRoles(IOrganizationService service, Guid userId)
        {
            List<string> userRoles = new List<string>();

            QueryExpression query = new QueryExpression("systemuserroles");
            query.ColumnSet = new ColumnSet("roleid");
            query.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);

            // Execute the query.
            EntityCollection result = service.RetrieveMultiple(query);

            // Process the results.
            foreach (Entity entity in result.Entities)
            {
                Guid roleId = entity.GetAttributeValue<Guid>("roleid");
                Entity role = service.Retrieve("role", roleId, new ColumnSet("name"));
                string roleName = role.GetAttributeValue<string>("name");

                userRoles.Add(roleName);
            }

            return userRoles;
        }
    }
}
