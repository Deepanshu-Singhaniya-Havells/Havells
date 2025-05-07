using System;
using System.Xml.Serialization;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using Microsoft.Xrm.Sdk;
using System.Web.Services.Description;
using System.Web.UI.WebControls;
using Microsoft.Xrm.Sdk.Organization;
using System.Web;
using Microsoft.Xrm.Sdk.Query;

namespace CreateFieldsFromExcelToD365
{
    public class ReadXML
    {
        static Dictionary<string, string> SDKMessageId = new Dictionary<string, string>();
        static Dictionary<string, string> viewIdsDict = new Dictionary<string, string>();
        static Dictionary<string, string> PluginTypeIdDict = new Dictionary<string, string>();
        public static void mainFunction(IOrganizationService service, IOrganizationService serviceSer)
        {
            changePluginTypeIDinSolution(service, serviceSer);

        }
        static void changePluginTypeIDinSolution(IOrganizationService service, IOrganizationService serviceSer)
        {
            string path = @"C:\Users\35405\OneDrive - Havells\Downloads\";
            string solutionName = @"\Havells_Plugin_10_0\";
            string customizations = GetXmlString(path + solutionName + "customizations.xml");
            string solution = GetXmlString(path + solutionName + "solution.xml");
            //CheangePluginTypeIDinCustomization(customizations, service, serviceSer);
            CheangeSDKMessageIDinCustomization(customizations, service, serviceSer);
            //foreach (string key in PluginTypeIdDict.Keys)
            //{
            //    // if (key == "{f3940bd5-684a-4a2a-a409-185d0d36f530}".ToUpper())
            //    Console.WriteLine(key + " : " + PluginTypeIdDict[key]);
            //    customizations = customizations.Replace(key.ToLower(), PluginTypeIdDict[key].ToLower());
            //    solution = solution.Replace(key.ToUpper(), PluginTypeIdDict[key].ToUpper());
            //}
            foreach (string key in SDKMessageId.Keys)
            {
                // if (key == "{f3940bd5-684a-4a2a-a409-185d0d36f530}".ToUpper())
                Console.WriteLine(key + " : " + SDKMessageId[key]);
                customizations = customizations.Replace(key.ToLower(), SDKMessageId[key].ToLower());
                solution = solution.Replace(key.ToUpper(), SDKMessageId[key].ToUpper());
            }


            Console.WriteLine("dione");

            foreach (string key in SDKMessageId.Keys)
            {
                // if (key == "{f3940bd5-684a-4a2a-a409-185d0d36f530}".ToUpper())
                Console.WriteLine(key + " : " + SDKMessageId[key]);
                //customizations = customizations.Replace(key.ToLower(), SDKMessageId[key].ToLower());
                //solution = solution.Replace(key.ToUpper(), SDKMessageId[key].ToUpper());
            }
            Console.WriteLine("dione");
        }

        static public Guid GetSdkMessageId(IOrganizationService serviceSer, IOrganizationService service, string SdkMessageID)
        {
            try
            {
                Entity entity = service.Retrieve("sdkmessage", new Guid(SdkMessageID), new ColumnSet(true));

                string SdkMessageName = entity.GetAttributeValue<string>("name");
                Console.WriteLine(SdkMessageName);
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
        static public Guid GetSdkMessageFilterId(IOrganizationService serviceSer, string EntityLogicalName, Guid sdkMessageId)
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
        static void CheangeSDKMessageIDinCustomization(string customizations, IOrganizationService service, IOrganizationService serviceSer)
        {
            customizations = customizations.Replace("<SdkMessageId>", "!");
            customizations = customizations.Replace("</SdkMessageId>", "$");
            string[] viewIds = customizations.Split('!');

            foreach (string viewIdString in viewIds)
            {
                string[] _viewidArray = viewIdString.Split('$');
                if (_viewidArray.Length == 2)
                {
                    string viewidorignal = _viewidArray[0];
                    var dictval = from x in SDKMessageId
                                  where x.Key.Contains(viewidorignal)
                                  select x;
                    if (dictval.ToList().Count == 0)
                    {
                        Guid newViewId = GetSdkMessageId(service, serviceSer, viewidorignal);
                        if (newViewId == null)
                        {
                            Console.WriteLine("ererer");
                        }
                        SDKMessageId.Add(viewidorignal, newViewId.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("fffff");
                }
            }
            Console.WriteLine();
        }



        static void CheangePluginTypeIDinCustomization(string customizations, IOrganizationService service, IOrganizationService serviceSer)
        {
            customizations = customizations.Replace("<PluginTypeId>", "@");
            customizations = customizations.Replace("</PluginTypeId>", "^");
            string[] viewIds = customizations.Split('@');

            foreach (string viewIdString in viewIds)
            {
                string[] _viewidArray = viewIdString.Split('^');
                if (_viewidArray.Length == 2)
                {
                    string viewidorignal = _viewidArray[0];
                    var dictval = from x in PluginTypeIdDict
                                  where x.Key.Contains(viewidorignal)
                                  select x;
                    if (dictval.ToList().Count == 0)
                    {
                        Guid newViewId = RetrivePlugins(service, serviceSer, viewidorignal);
                        if (newViewId == null)
                        {
                            Console.WriteLine("ererer");
                        }
                        PluginTypeIdDict.Add(viewidorignal, newViewId.ToString());
                    }
                }
                else
                {
                        Console.WriteLine("fffff");
                }

            }


            Console.WriteLine();
        }


        public static Guid RetrivePlugins(IOrganizationService service, IOrganizationService serviceSer, string plugintypeId)
        {
            Guid pluginId = Guid.Empty;
            try
            {
                Entity pluginTypeEntity = service.Retrieve("plugintype", new Guid(plugintypeId), new ColumnSet(true));
                Console.WriteLine("Plugins are creating");
                var query = new QueryExpression("plugintype");
                query.Criteria.AddCondition("name", ConditionOperator.Equal, pluginTypeEntity.GetAttributeValue<string>("name"));
                query.Criteria.AddCondition("assemblyname", ConditionOperator.Equal, "Havells_Plugin");
                query.ColumnSet = new ColumnSet(true);
                var results = serviceSer.RetrieveMultiple(query);
                Console.WriteLine("Total Plugins is " + results.Entities.Count);
                int done = 1;
                if (results.Entities.Count == 1)
                    pluginId = results[0].Id;
                else
                {
                    pluginId = Guid.Empty;

                    Console.WriteLine("done");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return pluginId;
        }
        static void changeViewIdinSolution(IOrganizationService service, IOrganizationService serviceSer)
        {
            string path = @"C:\Users\35405\OneDrive - Havells\Downloads\";
            string solutionName = @"Solution25\";
            string customizations = GetXmlString(path + solutionName + "customizations.xml");
            string solution = GetXmlString(path + solutionName + "solution.xml");
            CheangeViewIDinCustomization(customizations, service, serviceSer);

            foreach (string key in viewIdsDict.Keys)
            {
                // if (key == "{f3940bd5-684a-4a2a-a409-185d0d36f530}".ToUpper())
                Console.WriteLine(key + " : " + viewIdsDict[key]);
                customizations = customizations.Replace(key.ToLower(), viewIdsDict[key].ToLower());
                solution = solution.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
            }
            Console.WriteLine("dione");
        }
        static void CheangeViewIDinCustomization(string customizations, IOrganizationService service, IOrganizationService serviceSer)
        {
            customizations = customizations.Replace("<savedqueryid>", "@");
            customizations = customizations.Replace("</savedqueryid>", "#");
            string[] viewIds = customizations.Split('@');
            foreach (string viewIdString in viewIds)
            {
                string[] _viewidArray = viewIdString.Split('#');
                if (_viewidArray[0][0] == '{')
                {
                    string viewidorignal = _viewidArray[0];
                    var dictval = from x in viewIdsDict
                                  where x.Key.Contains(viewidorignal)
                                  select x;
                    if (dictval.ToList().Count == 0)
                    {
                        string newViewId = FormModfication.getDemoViewId(viewidorignal, service, serviceSer);
                        if (newViewId == null)
                        {
                            Console.WriteLine("ererer");
                        }
                        newViewId = newViewId == null ? "00000000-0000-0000-0000-000000000000" : newViewId;
                        viewIdsDict.Add(viewidorignal, "{" + newViewId + "}");
                    }
                }
            }


            Console.WriteLine();
        }

        static string GetXmlString(string strFile)
        {
            // Load the xml file into XmlDocument object.
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(strFile);
            }
            catch (XmlException e)
            {
                Console.WriteLine(e.Message);
            }
            // Now create StringWriter object to get data from xml document.
            StringWriter sw = new StringWriter();
            XmlTextWriter xw = new XmlTextWriter(sw);
            xmlDoc.WriteTo(xw);
            return sw.ToString();
        }
        static void readXMLFile()
        {
            using (XmlReader reader = XmlReader.Create(@"C:\Users\35405\Downloads\ModelDrivenApps_1_0_0_0\customizations.xml"))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        //return only when you have START tag  
                        switch (reader.Name.ToString())
                        {
                            case "Name":
                                Console.WriteLine("Name of the Element is : " + reader.ReadString());
                                break;
                            case "Location":
                                Console.WriteLine("Your Location is : " + reader.ReadString());
                                break;
                        }
                    }
                    Console.WriteLine("");
                }
            }
        }
    }
    public class SerializeDeserialize<T>
    {
        StringBuilder sbData;
        StringWriter swWriter;
        XmlDocument xDoc;
        XmlNodeReader xNodeReader;
        XmlSerializer xmlSerializer;
        public SerializeDeserialize()
        {
            sbData = new StringBuilder();
        }
        public string SerializeData(T data)
        {
            XmlSerializer employeeSerializer = new XmlSerializer(typeof(T));
            swWriter = new StringWriter(sbData);
            employeeSerializer.Serialize(swWriter, data);
            return sbData.ToString();
        }
        public T DeserializeData(string dataXML)
        {
            xDoc = new XmlDocument();
            xDoc.LoadXml(dataXML);
            xNodeReader = new XmlNodeReader(xDoc.DocumentElement);
            xmlSerializer = new XmlSerializer(typeof(T));
            var employeeData = xmlSerializer.Deserialize(xNodeReader);
            T deserializedEmployee = (T)employeeData;
            return deserializedEmployee;
        }
    }


    [XmlRoot(ElementName = "AppModuleComponent")]
    public class AppModuleComponent
    {
        [XmlAttribute(AttributeName = "type")]
        public string Type { get; set; }
        [XmlAttribute(AttributeName = "schemaName")]
        public string SchemaName { get; set; }
        [XmlAttribute(AttributeName = "id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "AppModuleComponents")]
    public class AppModuleComponents
    {
        [XmlElement(ElementName = "AppModuleComponent")]
        public List<AppModuleComponent> AppModuleComponent { get; set; }
    }

    [XmlRoot(ElementName = "Role")]
    public class Role
    {
        [XmlAttribute(AttributeName = "id")]
        public string Id { get; set; }
    }

    [XmlRoot(ElementName = "AppModuleRoleMaps")]
    public class AppModuleRoleMaps
    {
        [XmlElement(ElementName = "Role")]
        public List<Role> Role { get; set; }
    }

    [XmlRoot(ElementName = "LocalizedName")]
    public class LocalizedName
    {
        [XmlAttribute(AttributeName = "description")]
        public string Description { get; set; }
        [XmlAttribute(AttributeName = "languagecode")]
        public string Languagecode { get; set; }
    }

    [XmlRoot(ElementName = "LocalizedNames")]
    public class LocalizedNames
    {
        [XmlElement(ElementName = "LocalizedName")]
        public LocalizedName LocalizedName { get; set; }
    }

    [XmlRoot(ElementName = "Description")]
    public class Description
    {
        [XmlAttribute(AttributeName = "description")]
        public string Description1 { get; set; }
        [XmlAttribute(AttributeName = "languagecode")]
        public string Languagecode { get; set; }
    }

    [XmlRoot(ElementName = "Descriptions")]
    public class Descriptions
    {
        [XmlElement(ElementName = "Description")]
        public Description Description { get; set; }
    }

    [XmlRoot(ElementName = "AppModule")]
    public class AppModule
    {
        [XmlElement(ElementName = "UniqueName")]
        public string UniqueName { get; set; }
        [XmlElement(ElementName = "IntroducedVersion")]
        public string IntroducedVersion { get; set; }
        [XmlElement(ElementName = "WebResourceId")]
        public string WebResourceId { get; set; }
        [XmlElement(ElementName = "OptimizedFor")]
        public string OptimizedFor { get; set; }
        [XmlElement(ElementName = "statecode")]
        public string Statecode { get; set; }
        [XmlElement(ElementName = "statuscode")]
        public string Statuscode { get; set; }
        [XmlElement(ElementName = "FormFactor")]
        public string FormFactor { get; set; }
        [XmlElement(ElementName = "ClientType")]
        public string ClientType { get; set; }
        [XmlElement(ElementName = "NavigationType")]
        public string NavigationType { get; set; }
        [XmlElement(ElementName = "AppModuleComponents")]
        public AppModuleComponents AppModuleComponents { get; set; }
        [XmlElement(ElementName = "AppModuleRoleMaps")]
        public AppModuleRoleMaps AppModuleRoleMaps { get; set; }
        [XmlElement(ElementName = "LocalizedNames")]
        public LocalizedNames LocalizedNames { get; set; }
        [XmlElement(ElementName = "Descriptions")]
        public Descriptions Descriptions { get; set; }
    }

    [XmlRoot(ElementName = "AppModules")]
    public class AppModules
    {
        [XmlElement(ElementName = "AppModule")]
        public List<AppModule> AppModule { get; set; }
    }

    [XmlRoot(ElementName = "Languages")]
    public class Languages
    {
        [XmlElement(ElementName = "Language")]
        public string Language { get; set; }
    }

    [XmlRoot(ElementName = "ImportExportXml")]
    public class ImportExportXml
    {
        [XmlElement(ElementName = "Entities")]
        public string Entities { get; set; }
        [XmlElement(ElementName = "Roles")]
        public string Roles { get; set; }
        [XmlElement(ElementName = "Workflows")]
        public string Workflows { get; set; }
        [XmlElement(ElementName = "FieldSecurityProfiles")]
        public string FieldSecurityProfiles { get; set; }
        [XmlElement(ElementName = "Templates")]
        public string Templates { get; set; }
        [XmlElement(ElementName = "EntityMaps")]
        public string EntityMaps { get; set; }
        [XmlElement(ElementName = "EntityRelationships")]
        public string EntityRelationships { get; set; }
        [XmlElement(ElementName = "OrganizationSettings")]
        public string OrganizationSettings { get; set; }
        [XmlElement(ElementName = "optionsets")]
        public string Optionsets { get; set; }
        [XmlElement(ElementName = "CustomControls")]
        public string CustomControls { get; set; }
        [XmlElement(ElementName = "AppModules")]
        public AppModules AppModules { get; set; }
        [XmlElement(ElementName = "EntityDataProviders")]
        public string EntityDataProviders { get; set; }
        [XmlElement(ElementName = "Languages")]
        public Languages Languages { get; set; }
        [XmlAttribute(AttributeName = "xsi", Namespace = "http://www.w3.org/2000/xmlns/")]
        public string Xsi { get; set; }
    }
}
