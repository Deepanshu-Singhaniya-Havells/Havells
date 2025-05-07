using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace AE01.Assign_Roles
{
    internal class AssignRoles
    {
        private readonly IOrganizationService service;

        public AssignRoles(IOrganizationService _service)
        {
            this.service = _service;
        }
        internal void AssignInventoryUserRole()
        {
            // Assign roles to users 
            string fetchXml = @"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='hil_claimheader'>
                <attribute name='ownerid' />
                <order attribute='ownerid' descending='false' />
                <filter type='and'>
                <condition attribute='statecode' operator='eq' value='0' />
                <condition attribute='hil_fiscalmonth' operator='in'>
                <value uiname='202405' uitype='hil_claimperiod'>{56DE3041-22CE-EE11-904C-000D3A3E3D4E}</value>
                <value uiname='202406' uitype='hil_claimperiod'>{58DE3041-22CE-EE11-904C-000D3A3E3D4E}</value>
                <value uiname='202407' uitype='hil_claimperiod'>{5ADE3041-22CE-EE11-904C-000D3A3E3D4E}</value>
                <value uiname='202408' uitype='hil_claimperiod'>{5CDE3041-22CE-EE11-904C-000D3A3E3D4E}</value>
                </condition>
                </filter>
                <link-entity name='systemuser' from='systemuserid' to='owninguser' link-type='inner' alias='ac'>
                <filter type='and'>
                <condition attribute='positionid' operator='eq' value='{4A1AA189-1208-E911-A94D-000D3AF0694E}' />
                <condition attribute='isdisabled' operator='eq' value='0' />
                </filter>
                </link-entity>
                </entity>
                </fetch>";

            // Execute the FetchXML query
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            int i = 1;
            // Assign the role to each user
            foreach (var entity in result.Entities)
            {
                if (entity.Contains("ownerid"))
                {
                    EntityReference owner = (EntityReference)entity["ownerid"];
                    if (AssignRoleToUser(owner.Id, "Havells Inventory"))
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"{i}| Assigned role to user with userId {owner.Id}");
                        EnableSpareInventory(owner.Id);
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"{i}| Not able to assign role to the user with userid {owner.Id}");
                    }
                }
                i++;
            }
            Console.WriteLine("Role 'Havells Inventory' assigned to all users.");
        }

        internal void AssignInventoryApproverRole()
        {
            string fetchXml = @"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='systemuser'>
                    <attribute name='fullname' />
                    <attribute name='businessunitid' />
                    <attribute name='title' />
                    <attribute name='address1_telephone1' />
                    <attribute name='positionid' />
                    <attribute name='systemuserid' />
                    <order attribute='fullname' descending='false' />
                    <filter type='and'>
                      <condition attribute='positionid' operator='in'>
                        <value uiname='ASH' uitype='position'>{291AA189-1208-E911-A94D-000D3AF0694E}</value>
                        <value uiname='BSH' uitype='position'>{CEB72575-1208-E911-A94D-000D3AF0694E}</value>
                        <value uiname='CCO' uitype='position'>{B9A4A87F-1208-E911-A94D-000D3AF0694E}</value>
                      </condition>
                      <condition attribute='isdisabled' operator='eq' value='0' />
                    </filter>
                  </entity>
                </fetch>";

            EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            int i = 1;

            foreach (var entity in result.Entities)
            {
                if (AssignApproverRoleToUser(entity.Id))
                {
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"{i}| Assigned role to user with userId {entity.Id}");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"{i}| Not able to assign role to the user with userid {entity.Id}");
                }

                i++;

            }

        }

        bool AssignApproverRoleToUser(Guid userId)
        {
            try
            {
                // Retrieve the user's Business Unit
                Entity user = service.Retrieve("systemuser", userId, new ColumnSet("businessunitid"));
                Guid businessUnitId = user.GetAttributeValue<EntityReference>("businessunitid").Id;

                // Get the role ID for "Havells Inventory" in the user's Business Unit
                QueryExpression roleQuery = new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("name", ConditionOperator.Equal, "Havells Inventory (Approver)"),
                            new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId)
                        }
                    }
                };
                Entity role = service.RetrieveMultiple(roleQuery).Entities.FirstOrDefault();
                if (role == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Role 'Havells Inventory' not found for Business Unit {businessUnitId}.");
                    return false;
                }
                Guid roleId = role.Id;

                // Check if the user already has the role
                QueryExpression userRoleQuery = new QueryExpression("systemuserroles")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, userId),
                            new ConditionExpression("roleid", ConditionOperator.Equal, roleId)
                        }
                    } 
                };
                Entity existingUserRole = service.RetrieveMultiple(userRoleQuery).Entities.FirstOrDefault();
                if (existingUserRole != null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"User {userId} already has the role 'Havells Inventory Approver'.");
                    return true;
                }

                // Create a new association between the user and the role
                service.Associate(
                    "systemuser",
                    userId,
                    new Relationship("systemuserroles_association"),
                    new EntityReferenceCollection { new EntityReference("role", roleId) }
                );

                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error assigning role to user: " + ex.Message);
                return false;
            }
        }

        void EnableSpareInventory(Guid userId)
        {


            Entity user = service.Retrieve("systemuser", userId, new ColumnSet("hil_account"));
            if (user.Contains("hil_account"))
            {
                Entity account = service.Retrieve("account", user.GetAttributeValue<EntityReference>("hil_account").Id, new ColumnSet("hil_spareinventoryenabled"));
                if (account.GetAttributeValue<bool>("hil_spareinventoryenabled") == false)
                {
                    account["hil_spareinventoryenabled"] = true;
                    service.Update(account);
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine("Enabled spare part inventory for account with id " + account.Id);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Spare Inventory already enabled for the user's account.");
                }

            }

        }

        internal bool AssignRoleToUser(Guid userId, string roleName)
        {
            try
            {
                // Retrieve the user's Business Unit
                Entity user = service.Retrieve("systemuser", userId, new ColumnSet("businessunitid"));
                Guid businessUnitId = user.GetAttributeValue<EntityReference>("businessunitid").Id;

                // Get the role ID for Security Role in the user's Business Unit
                QueryExpression roleQuery = new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("name", ConditionOperator.Equal, roleName),
                            new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId)
                        }
                    }
                };
                Entity role = service.RetrieveMultiple(roleQuery).Entities.FirstOrDefault();
                if (role == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Role '{roleName}' not found for Business Unit {businessUnitId}.");
                    return false;
                }
                Guid roleId = role.Id;

                // Check if the user already has the role
                QueryExpression userRoleQuery = new QueryExpression("systemuserroles")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, userId),
                            new ConditionExpression("roleid", ConditionOperator.Equal, roleId)
                        }
                    }
                };
                Entity existingUserRole = service.RetrieveMultiple(userRoleQuery).Entities.FirstOrDefault();
                if (existingUserRole != null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"User {userId} already has the role '{roleName}'.");
                    return true;
                }

                // Create a new association between the user and the role
                service.Associate(
                    "systemuser",
                    userId,
                    new Relationship("systemuserroles_association"),
                    new EntityReferenceCollection { new EntityReference("role", roleId) }
                );

                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error assigning role to user: " + ex.Message);
                return false;
            }
        }

        internal void PrintAllRoles()
        {
            try
            {
                // Query to retrieve all roles
                QueryExpression query = new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("name", "roleid"),
                };

                // Retrieve roles
                EntityCollection roles = service.RetrieveMultiple(query);

                // Print roles
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("List of all roles:");
                foreach (var role in roles.Entities)
                {
                    string roleName = role.GetAttributeValue<string>("name");
                    Guid roleId = role.Id;
                    Console.WriteLine($"Role Name: {roleName}, Role ID: {roleId}");
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error retrieving roles: " + ex.Message);
            }
        }

        internal bool RemoveRoleFromUser(Guid userId, string roleName)
        {
            try
            {
                // Retrieve the user's Business Unit
                Entity user = service.Retrieve("systemuser", userId, new ColumnSet("businessunitid"));
                Guid businessUnitId = user.GetAttributeValue<EntityReference>("businessunitid").Id;

                // Get the role ID for the specified role in the user's Business Unit
                QueryExpression roleQuery = new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                {
                    new ConditionExpression("name", ConditionOperator.Equal, roleName),
                    new ConditionExpression("businessunitid", ConditionOperator.Equal, businessUnitId)
                }
                    }
                };
                Entity role = service.RetrieveMultiple(roleQuery).Entities.FirstOrDefault();
                if (role == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Role '{roleName}' not found for Business Unit {businessUnitId}.");
                    return false;
                }
                Guid roleId = role.Id;

                // Check if the user already has the role
                QueryExpression userRoleQuery = new QueryExpression("systemuserroles")
                {
                    ColumnSet = new ColumnSet("roleid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                {
                    new ConditionExpression("systemuserid", ConditionOperator.Equal, userId),
                    new ConditionExpression("roleid", ConditionOperator.Equal, roleId)
                }
                    }
                };
                Entity existingUserRole = service.RetrieveMultiple(userRoleQuery).Entities.FirstOrDefault();
                if (existingUserRole == null)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"User {userId} does not have the role '{roleName}'.");
                    return true;
                }

                // Remove the association between the user and the role
                service.Disassociate(
                    "systemuser",
                    userId,
                    new Relationship("systemuserroles_association"),
                    new EntityReferenceCollection { new EntityReference("role", roleId) }
                );

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine($"Role '{roleName}' removed from user {userId}.");
                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error removing role from user: " + ex.Message);
                return false;
            }
        }

        public void GetUsersForCallMasking()
        {
            QueryExpression query = new QueryExpression("systemuser");
            query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("4a1aa189-1208-e911-a94d-000d3af0694e"));
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, new DateTime(2025,03,30));
            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);

            EntityCollection usersCollection = service.RetrieveMultiple(query);

            int count = usersCollection.Entities.Count;
            Console.WriteLine("The total users are: " + count);

            for (int i = 0; i < count; i++) {
                Console.WriteLine("Processing the user number: " + (i + 1));
                if(AssignRoleToUser(usersCollection.Entities[i].Id, "Franchise Call Masking"))
                {
                    Console.WriteLine($"Franchise Call maskign has been assigned to the user: {usersCollection.Entities[i].Id}");
                }

                Console.WriteLine();
            }


        }
    }
}
