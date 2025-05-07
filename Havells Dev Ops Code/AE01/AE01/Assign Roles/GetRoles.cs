using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AE01.Assign_Roles
{

    internal static class GetRoles
    {

        private static readonly List<string> allowedRoles = new List<string>() { "System Administrator", "Grievance MDM" };
        public static void ToTest(IOrganizationService service)
        {

            List<string> userRoles = GetUserRoles(service, new Guid("42110c68-fcee-e811-a949-000d3af03089"));

            if (!userRoles.Any(role => allowedRoles.Contains(role)))
            {
                Console.WriteLine("You are not authorized to re-assign the activity");
            }

        }


        public static List<string> GetUserRoles(IOrganizationService service, Guid userId)
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
