using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class RemoveSecurityRoles
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                int queryCount = 5000;
                int pageCount = 1;
                int recordsCount = 0;

                var _positions = new object[] { 
                    new Guid("4A1AA189-1208-E911-A94D-000D3AF0694E"), 
                    new Guid("0197EA9B-1208-E911-A94D-000D3AF0694E"), 
                    new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E"),
                    new Guid("DC9D4659-FDD7-EA11-A813-000D3AF05A4B"),
                    new Guid("4AC9C45C-1208-E911-A94D-000D3AF0694E"),
                    new Guid("5FEEFA64-1208-E911-A94D-000D3AF0694E"),
                    new Guid("6401626B-1208-E911-A94D-000D3AF0694E"),
                    new Guid("CEB72575-1208-E911-A94D-000D3AF0694E"),
                    new Guid("B9A4A87F-1208-E911-A94D-000D3AF0694E"),
                    new Guid("291AA189-1208-E911-A94D-000D3AF0694E"),
                    new Guid("50ADE124-241C-EB11-A813-0022486E3BAA"),
                    new Guid("08C2924A-241C-EB11-A813-0022486E3BAA")
                };
                string[] _roles = {
                    "Customer service app access",
                    "Project Service Automation app access",
                    "Sales Enterprise app access",
                    "DSE - FSM App",
                    "Sales Team Member",
                    "Salesperson",
                    "System Administrator",
                    "System Customizer",
                    "CEO-Business Manager",
                    "System Customizer - User Setup",
                    "Voice of the Customer app access role",
                    "Custom System Admin",
                    "Survey Feedback Publisher",
                    "FileStoreService App Access",
                    "National Aggregator ResQ",
                    "D365ChannelDefinitions Admin",
                    "D365ChannelDefinitions User",
                    "EAC App Access",
                    "EAC Reader App Access",
                    "System Customizer (Developer)",
                    "Warranty Administrator",
                    "XAccount Manager",
                    "XBranch Product Head (BPH)"
                };

                QueryExpression query = new QueryExpression("systemuser");
                query.ColumnSet = new ColumnSet("firstname", "lastname", "positionid", "domainname");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("isdisabled", ConditionOperator.Equal, false));
                query.Criteria.AddCondition(new ConditionExpression("accessmode", ConditionOperator.NotEqual,3));
                query.Criteria.AddCondition(new ConditionExpression("accessmode", ConditionOperator.NotEqual, 5));
                //query.Criteria.AddCondition(new ConditionExpression("positionid", ConditionOperator.In, new object[] { new Guid("4A1AA189-1208-E911-A94D-000D3AF0694E"), new Guid("0197EA9B-1208-E911-A94D-000D3AF0694E"), new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E")}));
                query.Criteria.AddCondition(new ConditionExpression("positionid", ConditionOperator.In, _positions));
                query.AddOrder("createdon", OrderType.Descending);

                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = queryCount;
                query.PageInfo.PageNumber = pageCount;

                EntityCollection userCollection = new EntityCollection();
                while (true)
                {
                    EntityCollection temp = _service.RetrieveMultiple(query);
                    if (temp.Entities != null)
                    {
                        recordsCount += temp.Entities.Count;
                        foreach (Entity entity in temp.Entities)
                        {
                            userCollection.Entities.Add(entity);
                        }
                    }
                    if (temp.MoreRecords)
                    {
                        // Increment the page number to retrieve the next page.
                        query.PageInfo.PageNumber++;

                        // Set the paging cookie to the paging cookie returned from current results.
                        query.PageInfo.PagingCookie = temp.PagingCookie;
                    }
                    else
                    {
                        break;
                    }
                }
                int rowCnt = 1;
                foreach (Entity entity in userCollection.Entities)
                {
                    string getUserRoles = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>\r\n  <entity name='role'>\r\n    <attribute name='name' />\r\n    <attribute name='businessunitid' />\r\n    <attribute name='roleid' />\r\n    <order attribute='name' descending='false' />\r\n    <link-entity name='systemuserroles' from='roleid' to='roleid' visible='false' intersect='true'>\r\n      <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='ab'>\r\n        <filter type='and'>\r\n          <condition attribute='systemuserid' operator='eq' uiname='CRM ADMIN' uitype='systemuser' value='{" + entity.Id + "}' />\r\n        </filter>\r\n      </link-entity>\r\n    </link-entity>\r\n  </entity>\r\n</fetch>";
                    EntityCollection sample = _service.RetrieveMultiple(new FetchExpression(getUserRoles));
                    Console.WriteLine("Processing... " + rowCnt++.ToString() + "/" + userCollection.Entities.Count + " User:  " + entity.GetAttributeValue<string>("domainname"));
                    string roleName = string.Empty;
                    string roleNameFilter = string.Empty;
                    foreach (var role in sample.Entities)
                    {
                        roleName = string.Empty;
                        //Role roleEntity = role.ToEntity<Role>();
                        roleName = role.GetAttributeValue<string>("name").Trim();
                        Console.WriteLine(" ----> Role Name: " + roleName);

                        roleNameFilter = Array.Find(_roles, e => e.Equals(roleName));

                        if (roleNameFilter != null)
                        {
                            if (roleNameFilter == "System Administrator" || roleName == "System Customizer" || roleName == "System Customizer - User Setup" || roleName == "System Customizer (Developer)")
                            {
                                Console.WriteLine("Confirm to Remove.....");
                                Console.ReadLine();
                                _service.Disassociate("systemuser", entity.Id, new Relationship("systemuserroles_association"), new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                            }
                            else
                            {
                                _service.Disassociate("systemuser", entity.Id, new Relationship("systemuserroles_association"), new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                            }
                            Console.WriteLine("Security Role Removed.");
                        }
                        else {
                            Console.WriteLine("Skip...");
                        }
                        #region Franchise DSE Technician Role Setup
                        //if (roleName.IndexOf("CCO") >= 0 || roleName == "Call Center (Home Advisory)" || roleName == "DSE - FSM App" || roleName == "Technician-(Y)" || roleName == "Franchisee-( Y)" || roleName == "Technician Updated - (Y)" || roleName == "Technician Updated -(Y)" || roleName == "Franchisee (New)" || roleName == "Franchisee(New)" || roleName == "Direct Engg-updated (Y)" || roleName == "Direct Engineer")
                        //    Console.WriteLine("Skip...");
                        //else
                        //{
                        //    if (roleName == "FileStoreService App Access" || roleName == "Voice of the Customer app access role" || roleName == "Customer service app access" || roleName == "Sales Enterprise app access" || roleName == "Project Service Automation app access")
                        //    {
                        //        _service.Disassociate("systemuser", entity.Id, new Relationship("systemuserroles_association"), new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                        //    }
                        //    else {
                        //        Console.WriteLine("Confirm to Remove.....");
                        //        Console.ReadLine();
                        //        _service.Disassociate("systemuser", entity.Id, new Relationship("systemuserroles_association"), new EntityReferenceCollection() { new EntityReference("role", role.Id) });
                        //    }
                        //    Console.WriteLine("Security Role Removed.");
                        //}
                        #endregion
                    }
                }
            }
        }
        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
