using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

// This file is used to gt 
namespace AE01.Call_Masking
{
    internal class AssignFCMRole
    {
        private IOrganizationService service;
        private Guid _userId = new Guid("6e2ccd4d-da97-ed11-aad1-6045bdac5778");
        private string _givenRole = "Franchise Call Masking";

        public AssignFCMRole(IOrganizationService _service)
        {
            service = _service;
        }
       
        public void AssignRole()
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "role",
                ColumnSet = new ColumnSet("roleid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                                {

                                    new ConditionExpression
                                    {
                                        AttributeName = "name",
                                        Operator = ConditionOperator.Equal,
                                        Values = {_givenRole}
                                    }
                    }
                }
            };

            EntityCollection roles = service.RetrieveMultiple(query);
            if (roles.Entities.Count > 0)
            {
                Entity callMaskingRole = roles.Entities[0];
                Console.WriteLine("Role {0} is retrieved for the role assignment.", _givenRole);
                Guid _roleId = callMaskingRole.Id;

                // Associate the user with the role.
                if (_roleId != Guid.Empty && _userId != Guid.Empty)
                {
                    service.Associate(
                                "systemuser",
                                _userId,
                                new Relationship("systemuserroles_association"),
                                new EntityReferenceCollection() { new EntityReference("role", _roleId) });

                    Console.WriteLine("Role is associated with the user.");
                }
            }
        }
    }
}
