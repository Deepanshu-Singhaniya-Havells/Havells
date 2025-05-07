using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AE01.Miscellaneous
{
    internal class SystemAnalytics
    {
        private readonly IOrganizationService service;

        public SystemAnalytics(IOrganizationService _service)
        {
            this.service = _service;
        }

        public void GetSystemAnalytics()
        {
            try
            {
                Console.WriteLine("Fetching system analytics...");

                // Step 1: Get active and inactive users
                var userAnalytics = GetUserAnalytics();
                Console.WriteLine($"Active Users: {userAnalytics.ActiveUsers}, Inactive Users: {userAnalytics.InactiveUsers}");

                // Step 2: Get plugin usage statistics
                var pluginUsage = GetPluginUsage();
                Console.WriteLine("Most Used Plugins:");
                foreach (var plugin in pluginUsage)
                {
                    Console.WriteLine($"Plugin: {plugin.Key}, Execution Count: {plugin.Value}");
                }

                // Step 3: Get table usage statistics
                var tableUsage = GetTableUsage();
                Console.WriteLine("Most Used Tables:");
                foreach (var table in tableUsage)
                {
                    Console.WriteLine($"Table: {table.Key}, Record Count: {table.Value}");
                }

                // Step 4: Get web resource usage
                var webResourceCount = GetWebResourceCount();
                Console.WriteLine($"Total Web Resources: {webResourceCount}");

                // Step 5: Get system jobs (e.g., workflows, async operations)
                var systemJobs = GetSystemJobAnalytics();
                Console.WriteLine($"Total System Jobs: {systemJobs.TotalJobs}, Completed: {systemJobs.CompletedJobs}, Failed: {systemJobs.FailedJobs}");

                Console.WriteLine("System analytics fetched successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching system analytics: {ex.Message}");
            }
        }


        public void GetEntityRelationShips()
        {

        }

        private (int ActiveUsers, int InactiveUsers) GetUserAnalytics()
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("isdisabled")
            };

            var users = service.RetrieveMultiple(query);
            int activeUsers = 0, inactiveUsers = 0;

            foreach (var user in users.Entities)
            {
                bool isDisabled = user.GetAttributeValue<bool>("isdisabled");
                if (isDisabled)
                    inactiveUsers++;
                else
                    activeUsers++;
            }

            return (activeUsers, inactiveUsers);
        }

        private Dictionary<string, int> GetPluginUsage()
        {
            var pluginUsage = new Dictionary<string, int>();

            var query = new QueryExpression("sdkmessageprocessingstep")
            {
                ColumnSet = new ColumnSet("name", "eventhandler")
            };

            var steps = service.RetrieveMultiple(query);

            foreach (var step in steps.Entities)
            {
                string pluginName = step.GetAttributeValue<EntityReference>("eventhandler")?.Name;
                if (pluginName != null)
                {
                    if (pluginUsage.ContainsKey(pluginName))
                        pluginUsage[pluginName]++;
                    else
                        pluginUsage[pluginName] = 1;
                }
            }

            return pluginUsage;
        }

        private Dictionary<string, int> GetTableUsage()
        {
            var tableUsage = new Dictionary<string, int>();

            var query = new QueryExpression("entity")
            {
                ColumnSet = new ColumnSet("logicalname", "objecttypecode")
            };

            var entities = service.RetrieveMultiple(query);

            foreach (var entity in entities.Entities)
            {
                string tableName = entity.GetAttributeValue<string>("logicalname");
                if (!string.IsNullOrEmpty(tableName))
                {
                    var countQuery = new QueryExpression(tableName)
                    {
                        ColumnSet = new ColumnSet(false)
                    };

                    var count = service.RetrieveMultiple(countQuery).Entities.Count;
                    tableUsage[tableName] = count;
                }
            }

            return tableUsage;
        }

        private int GetWebResourceCount()
        {
            var query = new QueryExpression("webresource")
            {
                ColumnSet = new ColumnSet("name")
            };

            var webResources = service.RetrieveMultiple(query);
            return webResources.Entities.Count;
        }

        private (int TotalJobs, int CompletedJobs, int FailedJobs) GetSystemJobAnalytics()
        {
            var query = new QueryExpression("asyncoperation")
            {
                ColumnSet = new ColumnSet("statecode", "statuscode")
            };

            var jobs = service.RetrieveMultiple(query);
            int totalJobs = jobs.Entities.Count;
            int completedJobs = 0, failedJobs = 0;

            foreach (var job in jobs.Entities)
            {
                int stateCode = job.GetAttributeValue<OptionSetValue>("statecode").Value;
                int statusCode = job.GetAttributeValue<OptionSetValue>("statuscode").Value;

                if (stateCode == 3) // Completed
                    completedJobs++;
                else if (statusCode == 31) // Failed
                    failedJobs++;
            }

            return (totalJobs, completedJobs, failedJobs);
        }
    }

}
