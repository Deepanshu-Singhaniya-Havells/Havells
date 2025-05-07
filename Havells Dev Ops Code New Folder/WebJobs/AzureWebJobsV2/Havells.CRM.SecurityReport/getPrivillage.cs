using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells.CRM.SecurityReport
{
    public class getPrivillage
    {
        public static void getProvillage(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = "role";
            query.ColumnSet = new ColumnSet("name");
            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, new Guid("574D775D-60EB-E811-A96C-000D3AF05828"));
            query.Criteria.AddCondition("roleid", ConditionOperator.Equal, new Guid("77FB7C3B-C37C-EB11-A812-0022486EAE0C"));
            LinkEntity systemUseRole = new LinkEntity();
            systemUseRole.LinkFromEntityName = "role";
            systemUseRole.LinkFromAttributeName = "roleid";
            systemUseRole.LinkToEntityName = "systemuserroles";
            systemUseRole.LinkToAttributeName = "roleid";
            systemUseRole.JoinOperator = JoinOperator.Inner;
            systemUseRole.EntityAlias = "SUR";

            //LinkEntity userRoles = new LinkEntity();
            //userRoles.LinkFromEntityName = "systemuserroles";
            //userRoles.LinkFromAttributeName = "systemuserid";
            //userRoles.LinkToEntityName = "systemuser";
            //userRoles.LinkToAttributeName = "systemuserid";
            //userRoles.JoinOperator = JoinOperator.Inner;
            //userRoles.EntityAlias = "SU";
            //userRoles.Columns = new ColumnSet("fullname");

            LinkEntity rolePrivileges = new LinkEntity();
            rolePrivileges.LinkFromEntityName = "role";
            rolePrivileges.LinkFromAttributeName = "roleid";
            rolePrivileges.LinkToEntityName = "roleprivileges";
            rolePrivileges.LinkToAttributeName = "roleid";
            rolePrivileges.JoinOperator = JoinOperator.Inner;
            rolePrivileges.EntityAlias = "RP";
            rolePrivileges.Columns = new ColumnSet("privilegedepthmask");

            LinkEntity privilege = new LinkEntity();
            privilege.LinkFromEntityName = "roleprivileges";
            privilege.LinkFromAttributeName = "privilegeid";
            privilege.LinkToEntityName = "privilege";
            privilege.LinkToAttributeName = "privilegeid";
            privilege.JoinOperator = JoinOperator.Inner;
            privilege.EntityAlias = "P";
            privilege.Columns = new ColumnSet("name", "accessright");

            LinkEntity privilegeObjectTypeCodes = new LinkEntity();
            privilegeObjectTypeCodes.LinkFromEntityName = "roleprivileges";
            privilegeObjectTypeCodes.LinkFromAttributeName = "privilegeid";
            privilegeObjectTypeCodes.LinkToEntityName = "privilegeobjecttypecodes";
            privilegeObjectTypeCodes.LinkToAttributeName = "privilegeid";
            privilegeObjectTypeCodes.JoinOperator = JoinOperator.Inner;
            privilegeObjectTypeCodes.EntityAlias = "POTC";
            privilegeObjectTypeCodes.Columns = new ColumnSet("objecttypecode");
            privilegeObjectTypeCodes.Orders.Add(new OrderExpression("objecttypecode", OrderType.Ascending));
            //ConditionExpression conditionExpression = new ConditionExpression();
            //conditionExpression.AttributeName = "systemuserid";
            //conditionExpression.Operator = ConditionOperator.Equal;
            //conditionExpression.Values.Add(gUserId);

            //userRoles.LinkCriteria = new FilterExpression();
            //userRoles.LinkCriteria.Conditions.Add(conditionExpression);

            //systemUseRole.LinkEntities.Add(userRoles);
            query.LinkEntities.Add(systemUseRole);

            rolePrivileges.LinkEntities.Add(privilege);
            rolePrivileges.LinkEntities.Add(privilegeObjectTypeCodes);
            query.LinkEntities.Add(rolePrivileges);


            EntityCollection retUserRoles = service.RetrieveMultiple(query);

            Console.WriteLine("Retrieved {0} records", retUserRoles.Entities.Count);
            foreach (Entity rur in retUserRoles.Entities)
            {
                string UserName = String.Empty;
                string SecurityRoleName = String.Empty;
                string PriviligeName = String.Empty;
                string AccessLevel = String.Empty;
                string SecurityLevel = String.Empty;
                string EntityName = String.Empty;


                SecurityRoleName = (string)rur["name"];
                EntityName = ((AliasedValue)(rur["POTC.objecttypecode"])).Value.ToString();
                PriviligeName = ((AliasedValue)(rur["P.name"])).Value.ToString();



                switch (((AliasedValue)(rur["P.accessright"])).Value.ToString())
                {
                    case "1":
                        AccessLevel = "READ";
                        break;

                    case "2":
                        AccessLevel = "WRITE";
                        break;

                    case "4":
                        AccessLevel = "APPEND";
                        break;

                    case "16":
                        AccessLevel = "APPENDTO";
                        break;

                    case "32":
                        AccessLevel = "CREATE";
                        break;

                    case "65536":
                        AccessLevel = "DELETE";
                        break;

                    case "262144":
                        AccessLevel = "SHARE";
                        break;

                    case "524288":
                        AccessLevel = "ASSIGN";
                        break;

                    default:
                        AccessLevel = "";
                        break;
                }



                switch (((AliasedValue)(rur["RP.privilegedepthmask"])).Value.ToString())
                {
                    case "1":
                        SecurityLevel = "User";
                        break;

                    case "2":
                        SecurityLevel = "Business Unit";
                        break;

                    case "4":
                        SecurityLevel = "Parent: Child Business Unit";
                        break;

                    case "8":
                        SecurityLevel = "Organisation";
                        break;

                    default:
                        SecurityLevel = "";
                        break;
                }

                //if (EntityName.Contains("hil_tenderproduct"))
                {
                    ///Console.WriteLine("User name:" + ((AliasedValue)rur["SU.fullname"]).Value);
                    Console.WriteLine("Entity name:" + EntityName);
                    Console.WriteLine("Security Role name:" + rur["name"]);
                    Console.WriteLine("Privilige name:" + ((AliasedValue)rur["P.name"]).Value);
                    Console.WriteLine("Access Right :" + AccessLevel);
                    Console.WriteLine("Security Level:" + SecurityLevel);

                }
            }
        }

    }
}
