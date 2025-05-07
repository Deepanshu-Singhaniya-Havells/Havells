using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class MigrateAssingmentMatrix : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        private static string stateID = string.Empty;
        static MigrateAssingmentMatrix()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
            stateID = ConfigurationManager.AppSettings["stateId"].ToString();
        }
        public static void _MigrateAssingmentMatrix(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            String entityName = "hil_assignmentmatrix";
            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
            string[] stateIDsss = stateID.Split(';');
            int pageNumber = 1;
            string fetchXML = string.Empty;
            int count = 0;
            int RecStateCount = 0;
            foreach (string s in stateIDsss)
            {
                pageNumber = 1;
                RecStateCount = 0;
                while (true)
                {
                    fetchXML = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""hil_assignmentmatrix"">
                                    <all-attributes />
                                    <order attribute=""hil_franchiseedirectengineer"" descending=""false"" />
                                    <filter type=""and"">
                                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                    </filter>
                                    <link-entity name=""hil_pincode"" from=""hil_pincodeid"" to=""hil_pincode"" link-type=""inner"" alias=""ag"">
                                      <filter type=""and"">
                                        <condition attribute=""hil_state"" operator=""eq"" value=""{s}"" />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
                    EntityCollection entCol = _servicePrd.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        try
                        {
                            int i = 1;
                            Console.WriteLine("Record Count " + entCol.Entities.Count);
                            #region foreachloop...
                            foreach (Entity entity1 in entCol.Entities)
                            {
                                i++;
                                if (getRecordExist(_serviceDev, entity1.LogicalName, entity1.Id))
                                    continue;
                                Console.Write(".");
                            Cont:
                                Entity _createEntity = new Entity(entity1.LogicalName, entity1.Id);
                                try
                                {
                                    var attributes = entity1.Attributes.Keys;
                                    foreach (string name in attributes)
                                    {
                                        if (name != "modifiedby" && name != "createdby" && name != "organizationid")
                                        {
                                            if (entity1[name].GetType().Name == "EntityReference")
                                            {
                                                if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                                {
                                                    _createEntity[name] = new EntityReference("systemuser", RetriveSystemUser(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name).Id));
                                                }
                                                else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "uom")
                                                {
                                                    _createEntity[name] = new EntityReference("uom", getBaseUoM(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name).Id));
                                                }
                                                else
                                                {
                                                    _createEntity[name] = new EntityReference(entity1.GetAttributeValue<EntityReference>(name).LogicalName,
                                                        CreateRecordIfNotExist(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name), systemAdminRef));
                                                }
                                            }
                                            else
                                            {
                                                _createEntity[name] = entity1[name];
                                            }
                                        }
                                    }
                                    _serviceDev.Create(_createEntity);
                                    count++;
                                    RecStateCount++;
                                    Console.WriteLine("Done with id " + _createEntity.Id + " Count " + count);
                                }
                                catch (Exception ex)
                                {
                                    if (ex.Message != "Cannot insert duplicate key.")
                                    {
                                        Console.WriteLine("Error " + ex.Message);
                                        if (ex.Message.Contains("Does Not Exist"))
                                        {
                                            string[] arr = ex.Message.Split(' ');
                                            string entityNameRef = arr[1].Replace("'", "");
                                            string entityIDRef = arr[5];
                                            migrateSingleRecord(_servicePrd, _serviceDev, entityNameRef, entityIDRef, systemAdminRef);
                                            goto Cont;
                                        }
                                        Console.WriteLine("Error " + ex.Message);
                                    }
                                }
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ERROR !!! " + ex.Message);
                        }
                        pageNumber++;
                    }
                    else
                    {
                        break;
                    }
                }
                Console.WriteLine("----------------------------------------------------------------------------------------------------------------------------");
                Console.WriteLine("State With ID " + s + " is Completed. Total Record Count "+ RecStateCount);
                Console.WriteLine("----------------------------------------------------------------------------------------------------------------------------");
            }
            Console.WriteLine("Count !!! " + count);
            Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
        }
        public static void AssingmentMatrixgetCount(IOrganizationService _servicePrd)
        {
            string[] StateIDss = {
                "1739dad8-bbf7-e811-a94c-000d3af06a98",
                "1f812cd7-bbf7-e811-a94c-000d3af06cd4",
                "1839dad8-bbf7-e811-a94c-000d3af06a98",
                "f018aed8-bbf7-e811-a94c-000d3af0677f",
                "21812cd7-bbf7-e811-a94c-000d3af06cd4",
                "1939dad8-bbf7-e811-a94c-000d3af06a98",
                "f118aed8-bbf7-e811-a94c-000d3af0677f",
                "f7c274de-bbf7-e811-a94c-000d3af0694e",
                "e268b5de-bbf7-e811-a94c-000d3af06cd4",
                "27fcb4df-bbf7-e811-a94c-000d3af06091",
                "e368b5de-bbf7-e811-a94c-000d3af06cd4",
                "5db02adc-bbf7-e811-a94c-000d3af06c56",
                "69027de0-bbf7-e811-a94c-000d3af0677f",
                "e468b5de-bbf7-e811-a94c-000d3af06cd4",
                "69ffa4e4-bbf7-e811-a94c-000d3af06c56",
                "e568b5de-bbf7-e811-a94c-000d3af06cd4",
                "78f313e6-bbf7-e811-a94c-000d3af0694e",
                "79f313e6-bbf7-e811-a94c-000d3af0694e",
                "6df1014a-bd9b-eb11-b1ac-0022486ec6c5",
                "2a1e1ae6-bbf7-e811-a94c-000d3af06091",
                "ad00c1e8-bbf7-e811-a94c-000d3af0677f",
                "2b1e1ae6-bbf7-e811-a94c-000d3af06091",
                "cfb59ce9-bbf7-e811-a94c-000d3af06cd4",
                "bef4fbe8-bbf7-e811-a94c-000d3af06a98",
                "2c1e1ae6-bbf7-e811-a94c-000d3af06091",
                "4ffc44ec-bbf7-e811-a94c-000d3af06091",
                "d0b59ce9-bbf7-e811-a94c-000d3af06cd4",
                "7fd358f0-bbf7-e811-a94c-000d3af0694e",
                "aeea8bef-bbf7-e811-a94c-000d3af06c56",
                "80d358f0-bbf7-e811-a94c-000d3af0694e",
                "5a167ff2-bbf7-e811-a94c-000d3af06091",
                "81d358f0-bbf7-e811-a94c-000d3af0694e",
                "afea8bef-bbf7-e811-a94c-000d3af06c56",
                "b0ea8bef-bbf7-e811-a94c-000d3af06c56",
                "5c167ff2-bbf7-e811-a94c-000d3af06091",
                "24bbbff5-bbf7-e811-a94c-000d3af06a98",
                "2b3266f8-bbf7-e811-a94c-000d3af06c56"
             };
            foreach (string s in StateIDss)
            {
                string fetchXML = string.Empty;
                int pageNumber = 1;
                int count = 0;
                while (true)
                {
                    fetchXML = $@"<fetch version=""1.0"" page=""{pageNumber}"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""hil_assignmentmatrix"">
                                    <all-attributes />
                                    <order attribute=""hil_franchiseedirectengineer"" descending=""false"" />
                                    <filter type=""and"">
                                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                    </filter>
                                    <link-entity name=""hil_pincode"" from=""hil_pincodeid"" to=""hil_pincode"" link-type=""inner"" alias=""ag"">
                                      <filter type=""and"">
                                        <condition attribute=""hil_state"" operator=""eq"" value=""{s}"" />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
                    EntityCollection entCol1 = _servicePrd.RetrieveMultiple(new FetchExpression(fetchXML));
                    count = count + entCol1.Entities.Count;
                    if (entCol1.Entities.Count > 0)
                    {
                        pageNumber++;
                    }
                    else
                    {
                        break;
                    }
                }
                Console.WriteLine(s + "|" + count);
            }
        }
    }
}
