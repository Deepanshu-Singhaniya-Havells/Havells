using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AE01.Miscellaneous
{
    internal class MetaData
    {
        private readonly IOrganizationService service;

        public MetaData(IOrganizationService _service)
        {
            this.service = _service;
        }

        public void ExportMetaData(string outputFilePath)
        {
            try
            {
                var environmentData = new Dictionary<string, object>();

                // Step 1: Retrieve all tables (entities) and their metadata
                Console.WriteLine("Retrieving table metadata...");
                var tables = GetAllTables();
                environmentData["Tables"] = tables;

                // Step 2: Retrieve all plugins and their registrations
                Console.WriteLine("Retrieving plugin registrations...");
                var plugins = GetAllPlugins();
                environmentData["Plugins"] = plugins;

                // Step 3: Retrieve all web resources
                Console.WriteLine("Retrieving web resources...");
                var webResources = GetAllWebResources();
                environmentData["WebResources"] = webResources;

                // Step 4: Serialize the data to JSON and save it to a file
                Console.WriteLine("Exporting metadata to file...");
                var json = JsonConvert.SerializeObject(environmentData, Formatting.Indented);
                File.WriteAllText(outputFilePath, json);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Environment metadata exported successfully to {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error exporting environment metadata: " + ex.Message);
            }
        }

        private List<object> GetAllTables()
        {
            var tables = new List<object>();

            // Retrieve all entities (tables) metadata
            var request = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.Entity | EntityFilters.Attributes,
                RetrieveAsIfPublished = true
            };

            var response = (RetrieveAllEntitiesResponse)service.Execute(request);

            foreach (var entityMetadata in response.EntityMetadata)
            {
                var table = new
                {
                    LogicalName = entityMetadata.LogicalName,
                    DisplayName = entityMetadata.DisplayName?.UserLocalizedLabel?.Label,
                    PrimaryIdAttribute = entityMetadata.PrimaryIdAttribute,
                    PrimaryNameAttribute = entityMetadata.PrimaryNameAttribute,
                    Attributes = GetAttributes(entityMetadata.Attributes)
                };

                tables.Add(table);
            }

            return tables;
        }

        private List<object> GetAttributes(AttributeMetadata[] attributes)
        {
            var attributeList = new List<object>();

            foreach (var attribute in attributes)
            {
                var attributeData = new
                {
                    LogicalName = attribute.LogicalName,
                    DisplayName = attribute.DisplayName?.UserLocalizedLabel?.Label,
                    AttributeType = attribute.AttributeTypeName?.Value
                };

                attributeList.Add(attributeData);
            }

            return attributeList;
        }

        private List<object> GetAllPlugins()
        {
            var plugins = new List<object>();

            // Query the plugin assembly and step registrations
            var query = new QueryExpression("sdkmessageprocessingstep")
            {
                ColumnSet = new ColumnSet("name", "eventhandler", "sdkmessageid", "stage", "mode", "rank"),
                LinkEntities =
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "sdkmessageprocessingstep",
                        LinkFromAttributeName = "sdkmessageid",
                        LinkToEntityName = "sdkmessage",
                        LinkToAttributeName = "sdkmessageid",
                        Columns = new ColumnSet("name"),
                        EntityAlias = "sdkmessage"
                    }
                }
            };

            var steps = service.RetrieveMultiple(query);

            foreach (var step in steps.Entities)
            {
                var plugin = new
                {
                    Name = step.GetAttributeValue<string>("name"),
                    EventHandler = step.GetAttributeValue<EntityReference>("eventhandler")?.Name,
                    Message = step.GetAttributeValue<AliasedValue>("sdkmessage.name")?.Value,
                    Stage = step.GetAttributeValue<OptionSetValue>("stage")?.Value,
                    Mode = step.GetAttributeValue<OptionSetValue>("mode")?.Value,
                    Rank = step.GetAttributeValue<int>("rank")
                };

                plugins.Add(plugin);
            }

            return plugins;
        }

        private List<object> GetAllWebResources()
        {
            var webResources = new List<object>();

            // Query the web resources
            var query = new QueryExpression("webresource")
            {
                ColumnSet = new ColumnSet("name", "displayname", "webresourcetype")
            };

            var resources = service.RetrieveMultiple(query);

            foreach (var resource in resources.Entities)
            {
                var webResource = new
                {
                    Name = resource.GetAttributeValue<string>("name"),
                    DisplayName = resource.GetAttributeValue<string>("displayname"),
                    ResourceType = resource.GetAttributeValue<OptionSetValue>("webresourcetype")?.Value
                };

                webResources.Add(webResource);
            }

            return webResources;
        }
    }
}



