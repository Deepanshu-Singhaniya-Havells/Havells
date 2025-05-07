using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.Helper
{
    public class HelperClass
    {
        public static readonly string _callMaskingRoleName = "Franchise Call Masking";
        public static string getUserPosition(Guid userid, IOrganizationService service, ITracingService tracingService)
        {
            tracingService.Trace("getUserPosition Started");
            string positionName = null;
            Entity user = service.Retrieve("systemuser", userid, new ColumnSet("positionid"));
            tracingService.Trace("User " + user.Id);
            if (user.Contains("positionid"))
            {
                tracingService.Trace(" UserPosition retived");

                positionName = user.GetAttributeValue<EntityReference>("positionid").Name;
                tracingService.Trace(" UserPosition " + positionName);
            }
            tracingService.Trace(" UserPosition not found");
            return positionName;
        }
        public static bool getUserSecurityRole(Guid userid, IOrganizationService service,string securityRole, ITracingService tracingService)
        {
            bool isRoleFound = false;
            QueryExpression qe = new QueryExpression("systemuserroles");
            qe.ColumnSet.AddColumns("systemuserid");
            qe.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userid);

            LinkEntity link1 = qe.AddLink("systemuser", "systemuserid", "systemuserid", JoinOperator.Inner);
            link1.Columns.AddColumns("fullname", "internalemailaddress");
            LinkEntity link = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            link.Columns.AddColumns("roleid", "name");
            EntityCollection results = service.RetrieveMultiple(qe);
            foreach (Entity Userrole in results.Entities)
            {
                if (Userrole.Attributes.Contains("role2.name"))
                {
                    tracingService.Trace("Role Name " + (Userrole.Attributes["role2.name"] as AliasedValue).Value.ToString());
                    if ((Userrole.Attributes["role2.name"] as AliasedValue).Value.ToString().ToLower() == securityRole.ToLower())//"Customer Asset Delink")
                        isRoleFound = true;
                }
            }
            return isRoleFound;
        }
    }
}
