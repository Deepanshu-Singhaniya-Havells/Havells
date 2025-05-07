using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.DependencyModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateFieldsFromExcelToD365
{
    public static class PluginMigration
    {
        public static void retivePluginAssemblly(IOrganizationService service, IOrganizationService serviceSer, string pluginName)
        {
            try
            {
                Console.WriteLine("Plugin Assembly Name " + pluginName);
                var query = new QueryExpression("pluginassembly");
                query.Criteria.AddCondition("name", ConditionOperator.Equal, pluginName);
                query.ColumnSet = new ColumnSet(true);
                var results = service.RetrieveMultiple(query);
                Entity entity = results[0];
                entity["ismanaged"] = false;
                entity["solutionid"] = new Guid();
                serviceSer.Create(entity);
                Console.WriteLine("@@@@@@@@@@@@ Plugin Assembly with name " + pluginName + " is created");
                RetrivePlugins(service, serviceSer, entity.Id);
                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("________________________________Error in Retriving " + pluginName + " Pluing is " + ex.Message);
            }
        }
        public static void RetrivePlugins(IOrganizationService service, IOrganizationService serviceSer, Guid pluignAssembly)
        {
            try
            {
                Console.WriteLine("Plugins are creating");
                var query = new QueryExpression("plugintype");
                query.Criteria.AddCondition("pluginassemblyid", ConditionOperator.Equal, pluignAssembly);
                query.ColumnSet = new ColumnSet(true);
                var results = service.RetrieveMultiple(query);
                Console.WriteLine("Total Plugins is " + results.Entities.Count);
                int done = 1;
                foreach (Entity entity in results.Entities)
                {

                    serviceSer.Create(entity);
                    Console.WriteLine("########## Pluginwith Name " + entity.GetAttributeValue<string>("name") + "is created");
                    CreatePluginSteps(service, serviceSer, entity.Id, pluignAssembly);
                    Console.WriteLine("Done Plugin " + done);
                    done++;
                }
                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("________________________________Error in Retriving " + pluignAssembly + " Pluing is " + ex.Message);
            }
        }
        public static void CreatePluginSteps(IOrganizationService service, IOrganizationService serviceSer, Guid pluginTypeId, Guid assemblyId)
        {
            try
            {
                Console.WriteLine("Plugins Steps are creating");
                var query = new QueryExpression("sdkmessageprocessingstep");
                query.Criteria.AddCondition("plugintypeid", ConditionOperator.Equal, pluginTypeId);
                query.ColumnSet = new ColumnSet(true);
                var results = service.RetrieveMultiple(query);
                foreach (Entity entity in results.Entities)
                {
                    Entity step = new Entity("sdkmessageprocessingstep");
                    step["name"] = entity.GetAttributeValue<string>("name");
                    step["description"] = entity.GetAttributeValue<string>("description");
                    step["configuration"] = entity.GetAttributeValue<string>("configuration");
                    step["mode"] = entity.GetAttributeValue<OptionSetValue>("mode");
                    step["rank"] = entity.GetAttributeValue<int>("rank");
                    step["stage"] = entity.GetAttributeValue<OptionSetValue>("stage");
                    step["supporteddeployment"] = entity.GetAttributeValue<OptionSetValue>("supporteddeployment");
                    step["invocationsource"] = entity.GetAttributeValue<OptionSetValue>("invocationsource");
                    if (entity.Contains("sdkmessagefilterid"))
                    {
                        Entity entity1 = service.Retrieve("sdkmessagefilter", entity.GetAttributeValue<EntityReference>("sdkmessagefilterid").Id, new ColumnSet("primaryobjecttypecode", "sdkmessageid"));
                        string objectTypeCode = entity1.GetAttributeValue<string>("primaryobjecttypecode");
                        Entity entity12 = service.Retrieve("sdkmessage", entity1.GetAttributeValue<EntityReference>("sdkmessageid").Id, new ColumnSet("name"));
                        string messageName = entity12.GetAttributeValue<string>("name");

                        Guid sdkMessageId = GetSdkMessageId(serviceSer, messageName);
                        Guid sdkMessageFilterId = GetSdkMessageFilterId(serviceSer, objectTypeCode, sdkMessageId);
                        step.Id = entity.Id;
                        step["plugintypeid"] = new EntityReference("plugintype", pluginTypeId);
                        step["sdkmessageid"] = new EntityReference("sdkmessage", sdkMessageId);
                        step["sdkmessagefilterid"] = new EntityReference("sdkmessagefilter", sdkMessageFilterId);

                        if (entity.Contains("impersonatinguserid"))
                        {
                            step["impersonatinguserid"] = new EntityReference("systemuser", RetriveSystemUser(service, serviceSer, entity.GetAttributeValue<EntityReference>("impersonatinguserid").Id));
                        }
                        try
                        {
                            Guid stepId = serviceSer.Create(step);
                            Console.WriteLine("!!!!!!!!!!! Plugin Step with Name " + entity.GetAttributeValue<string>("name") + "is created");
                            SdkMessageImage(service, serviceSer, stepId);
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Contains("Cannot insert duplicate key exception when executing non-query:"))
                                Console.WriteLine("Error in is " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("________________________________Error in Retriving " + pluginTypeId + " Pluing is " + ex.Message);
            }
        }
        private static Guid RetriveSystemUser(IOrganizationService service, IOrganizationService serviceSer, Guid user)
        {

            Entity prdUser = service.Retrieve("systemuser", user, new ColumnSet("domainname"));
            var query = new QueryExpression("systemuser");
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, prdUser.GetAttributeValue<string>("domainname"));
            query.ColumnSet = new ColumnSet(false);
            var results = service.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return Guid.Empty;
            else
                return results[0].Id;

        }
        private static Guid GetSdkMessageId(IOrganizationService serviceSer, string SdkMessageName)
        {
            try
            {
                //GET SDK MESSAGE QUERY
                QueryExpression sdkMessageQueryExpression = new QueryExpression("sdkmessage");
                sdkMessageQueryExpression.ColumnSet = new ColumnSet("sdkmessageid");
                sdkMessageQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression
                    {
                        AttributeName = "name",
                        Operator = ConditionOperator.Equal,
                        Values = {SdkMessageName}
                    },
                }
                };

                //RETRIEVE SDK MESSAGE
                EntityCollection sdkMessages = serviceSer.RetrieveMultiple(sdkMessageQueryExpression);
                if (sdkMessages.Entities.Count != 0)
                {
                    return sdkMessages.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK MessageName {0} was not found.", SdkMessageName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        static private Guid GetSdkMessageFilterId(IOrganizationService serviceSer, string EntityLogicalName, Guid sdkMessageId)
        {
            try
            {
                //GET SDK MESSAGE FILTER QUERY
                QueryExpression sdkMessageFilterQueryExpression = new QueryExpression("sdkmessagefilter");
                sdkMessageFilterQueryExpression.ColumnSet = new ColumnSet("sdkmessagefilterid");
                sdkMessageFilterQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression
                    {
                        AttributeName = "primaryobjecttypecode",
                        Operator = ConditionOperator.Equal,
                        Values = {EntityLogicalName}
                    },
                    new ConditionExpression
                    {
                        AttributeName = "sdkmessageid",
                        Operator = ConditionOperator.Equal,
                        Values = {sdkMessageId}
                    },
                }
                };

                //RETRIEVE SDK MESSAGE FILTER
                EntityCollection sdkMessageFilters = serviceSer.RetrieveMultiple(sdkMessageFilterQueryExpression);

                if (sdkMessageFilters.Entities.Count != 0)
                {
                    return sdkMessageFilters.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK Message Filter for {0} was not found.", EntityLogicalName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        static private void SdkMessageImage(IOrganizationService service, IOrganizationService serviceSer, Guid sdkmessageprocessingstepid)
        {
            try
            {
                //GET SDK MESSAGE FILTER QUERY
                QueryExpression sdkMessageImagesQueryExpression = new QueryExpression("sdkmessageprocessingstepimage");
                sdkMessageImagesQueryExpression.ColumnSet = new ColumnSet(true);
                sdkMessageImagesQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "sdkmessageprocessingstepid",
                            Operator = ConditionOperator.Equal,
                            Values = { sdkmessageprocessingstepid }
                        }
                    }
                };

                //RETRIEVE SDK MESSAGE FILTER
                EntityCollection sdkMessageImages = service.RetrieveMultiple(sdkMessageImagesQueryExpression);
                foreach (Entity image in sdkMessageImages.Entities)
                {
                    serviceSer.Create(image);
                    Console.WriteLine("%%%%%%%%%%%%%% Image is Created");
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
    }

    public static class PluginMigrationService
    {
        public static void retivePluginAssemblly(IOrganizationService serviceSer, string pluginName)
        {
            try
            {
                Console.WriteLine("Plugin Assembly Name " + pluginName);
                var query = new QueryExpression("pluginassembly");
                query.Criteria.AddCondition("name", ConditionOperator.Equal, pluginName);
                query.ColumnSet = new ColumnSet(true);
                var results = serviceSer.RetrieveMultiple(query);
                Entity entity = results[0];
                entity["ismanaged"] = false;
                entity["solutionid"] = new Guid();
                //serviceSer.Create(entity);
                Console.WriteLine("@@@@@@@@@@@@ Plugin Assembly with name " + pluginName + " is created");
                RetrivePlugins(serviceSer, entity.Id);
                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("________________________________Error in Retriving " + pluginName + " Pluing is " + ex.Message);
            }
        }
        public static void RetrivePlugins(IOrganizationService serviceSer, Guid pluignAssembly)
        {
            try
            {
                Console.WriteLine("Plugins are creating");
                var query = new QueryExpression("plugintype");
                query.Criteria.AddCondition("pluginassemblyid", ConditionOperator.Equal, pluignAssembly);
                query.ColumnSet = new ColumnSet(true);
                var results = serviceSer.RetrieveMultiple(query);
                Console.WriteLine("Total Plugins is " + results.Entities.Count);
                int done = 1;
                foreach (Entity entity in results.Entities)
                {

                    // serviceSer.Create(entity);
                    Console.WriteLine("########## Pluginwith Name " + entity.GetAttributeValue<string>("name") + "is created");

                    RetrieveDependentComponentsRequest dependentComponentsRequest =
                            new RetrieveDependentComponentsRequest
                            {
                                ComponentType = 90,
                                ObjectId = entity.Id
                            };
                    RetrieveDependentComponentsResponse dependentComponentsResponse = (RetrieveDependentComponentsResponse)serviceSer.Execute(dependentComponentsRequest);


                    //A more complete report requires more code
                    foreach (Entity d in dependentComponentsResponse.EntityCollection.Entities)
                    {
                        Console.WriteLine("########## Pluginwith Name " + entity.GetAttributeValue<string>("name") + "is created");


                        //  DependencyReport(d);
                    }

                    CreatePluginSteps(serviceSer, entity.Id, pluignAssembly);
                    Console.WriteLine("Done Plugin " + done);
                    serviceSer.Delete(entity.LogicalName, entity.Id);
                    done++;
                }
                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("________________________________Error in Retriving " + pluignAssembly + " Pluing is " + ex.Message);
            }
        }
        public static void CreatePluginSteps(IOrganizationService serviceSer, Guid pluginTypeId, Guid assemblyId)
        {
            try
            {
                Console.WriteLine("Plugins Steps are creating");
                var query = new QueryExpression("sdkmessageprocessingstep");
                query.Criteria.AddCondition("plugintypeid", ConditionOperator.Equal, pluginTypeId);
                query.ColumnSet = new ColumnSet(true);
                var results = serviceSer.RetrieveMultiple(query);
                foreach (Entity entity in results.Entities)
                {
                    Entity step = new Entity("sdkmessageprocessingstep");
                    step["name"] = entity.GetAttributeValue<string>("name");
                    step["description"] = entity.GetAttributeValue<string>("description");
                    step["configuration"] = entity.GetAttributeValue<string>("configuration");
                    step["mode"] = entity.GetAttributeValue<OptionSetValue>("mode");
                    step["rank"] = entity.GetAttributeValue<int>("rank");
                    step["stage"] = entity.GetAttributeValue<OptionSetValue>("stage");
                    step["supporteddeployment"] = entity.GetAttributeValue<OptionSetValue>("supporteddeployment");
                    step["invocationsource"] = entity.GetAttributeValue<OptionSetValue>("invocationsource");
                    if (entity.Contains("sdkmessagefilterid"))
                    {
                        Entity entity1 = serviceSer.Retrieve("sdkmessagefilter", entity.GetAttributeValue<EntityReference>("sdkmessagefilterid").Id, new ColumnSet("primaryobjecttypecode", "sdkmessageid"));
                        string objectTypeCode = entity1.GetAttributeValue<string>("primaryobjecttypecode");
                        Entity entity12 = serviceSer.Retrieve("sdkmessage", entity1.GetAttributeValue<EntityReference>("sdkmessageid").Id, new ColumnSet("name"));
                        string messageName = entity12.GetAttributeValue<string>("name");

                        Guid sdkMessageId = GetSdkMessageId(serviceSer, messageName);
                        Guid sdkMessageFilterId = GetSdkMessageFilterId(serviceSer, objectTypeCode, sdkMessageId);
                        step.Id = entity.Id;
                        step["plugintypeid"] = new EntityReference("plugintype", pluginTypeId);
                        step["sdkmessageid"] = new EntityReference("sdkmessage", sdkMessageId);
                        step["sdkmessagefilterid"] = new EntityReference("sdkmessagefilter", sdkMessageFilterId);

                        if (entity.Contains("impersonatinguserid"))
                        {
                            //step["impersonatinguserid"] = new EntityReference("systemuser", RetriveSystemUser(service, serviceSer, entity.GetAttributeValue<EntityReference>("impersonatinguserid").Id));
                        }
                        try
                        {
                            Guid stepId = entity.Id;// serviceSer.Create(step);
                            Console.WriteLine("!!!!!!!!!!! Plugin Step with Name " + entity.GetAttributeValue<string>("name") + "is created");
                            SdkMessageImage(serviceSer, stepId);
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Contains("Cannot insert duplicate key exception when executing non-query:"))
                                Console.WriteLine("Error in is " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("________________________________Error in Retriving " + pluginTypeId + " Pluing is " + ex.Message);
            }
        }
        private static Guid RetriveSystemUser(IOrganizationService service, IOrganizationService serviceSer, Guid user)
        {

            Entity prdUser = service.Retrieve("systemuser", user, new ColumnSet("domainname"));
            var query = new QueryExpression("systemuser");
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, prdUser.GetAttributeValue<string>("domainname"));
            query.ColumnSet = new ColumnSet(false);
            var results = service.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return Guid.Empty;
            else
                return results[0].Id;

        }
        private static Guid GetSdkMessageId(IOrganizationService serviceSer, string SdkMessageName)
        {
            try
            {
                //GET SDK MESSAGE QUERY
                QueryExpression sdkMessageQueryExpression = new QueryExpression("sdkmessage");
                sdkMessageQueryExpression.ColumnSet = new ColumnSet("sdkmessageid");
                sdkMessageQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression
                    {
                        AttributeName = "name",
                        Operator = ConditionOperator.Equal,
                        Values = {SdkMessageName}
                    },
                }
                };

                //RETRIEVE SDK MESSAGE
                EntityCollection sdkMessages = serviceSer.RetrieveMultiple(sdkMessageQueryExpression);
                if (sdkMessages.Entities.Count != 0)
                {
                    return sdkMessages.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK MessageName {0} was not found.", SdkMessageName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        static private Guid GetSdkMessageFilterId(IOrganizationService serviceSer, string EntityLogicalName, Guid sdkMessageId)
        {
            try
            {
                //GET SDK MESSAGE FILTER QUERY
                QueryExpression sdkMessageFilterQueryExpression = new QueryExpression("sdkmessagefilter");
                sdkMessageFilterQueryExpression.ColumnSet = new ColumnSet("sdkmessagefilterid");
                sdkMessageFilterQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression
                    {
                        AttributeName = "primaryobjecttypecode",
                        Operator = ConditionOperator.Equal,
                        Values = {EntityLogicalName}
                    },
                    new ConditionExpression
                    {
                        AttributeName = "sdkmessageid",
                        Operator = ConditionOperator.Equal,
                        Values = {sdkMessageId}
                    },
                }
                };

                //RETRIEVE SDK MESSAGE FILTER
                EntityCollection sdkMessageFilters = serviceSer.RetrieveMultiple(sdkMessageFilterQueryExpression);

                if (sdkMessageFilters.Entities.Count != 0)
                {
                    return sdkMessageFilters.Entities.First().Id;
                }
                throw new Exception(String.Format("SDK Message Filter for {0} was not found.", EntityLogicalName));
            }
            catch (InvalidPluginExecutionException invalidPluginExecutionException)
            {
                throw invalidPluginExecutionException;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        static private void SdkMessageImage(IOrganizationService serviceSer, Guid sdkmessageprocessingstepid)
        {
            try
            {
                //GET SDK MESSAGE FILTER QUERY
                QueryExpression sdkMessageImagesQueryExpression = new QueryExpression("sdkmessageprocessingstepimage");
                sdkMessageImagesQueryExpression.ColumnSet = new ColumnSet(true);
                sdkMessageImagesQueryExpression.Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "sdkmessageprocessingstepid",
                            Operator = ConditionOperator.Equal,
                            Values = { sdkmessageprocessingstepid }
                        }
                    }
                };

                //RETRIEVE SDK MESSAGE FILTER
                EntityCollection sdkMessageImages = serviceSer.RetrieveMultiple(sdkMessageImagesQueryExpression);
                foreach (Entity image in sdkMessageImages.Entities)
                {
                    serviceSer.Create(image);
                    Console.WriteLine("%%%%%%%%%%%%%% Image is Created");
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
    }
}
