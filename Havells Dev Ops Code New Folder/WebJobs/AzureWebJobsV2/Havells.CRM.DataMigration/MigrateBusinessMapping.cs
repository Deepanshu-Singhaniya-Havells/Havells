using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class MigrateBusinessMapping : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        private static Guid stateId = Guid.Empty;
        static MigrateBusinessMapping()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
            stateId = new Guid(ConfigurationManager.AppSettings["stateId"].ToString());
        }
        public static void UpdateMigratedBusMapping(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            string[] stateID =  {
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

            Parallel.ForEach(stateID, state =>
            {
                string entityName = "hil_businessmapping";
                //Console.WriteLine("****************************************** Entity " + entityName + " is started for State "+ s + ".****************************************** ");
                QueryExpression query = new QueryExpression(entityName.ToLower());
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, state);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int count = 0;
                int error = 0;
                int done = 0;
                try
                {

                    EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
                    string statename = entCol[0].GetAttributeValue<EntityReference>("hil_state").Name;
                    count = entCol.Entities.Count;
                    do
                    {
                        #region foreach loop...
                        foreach (Entity entity1 in entCol.Entities)
                        {
                           
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
                                _createEntity["createdby"] = systemAdminRef;
                                if (entity1.Contains("ownerid"))
                                    _createEntity["ownerid"] = systemAdminRef;
                                _createEntity["modifiedby"] = systemAdminRef;
                                string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
                                if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
                                    _serviceDev.Update(_createEntity);
                                done++;
                                Console.WriteLine("Done " + done + "/" + count + " Record of State " + statename);
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
                                    error++;
                                    Console.WriteLine("Error Count " + error);
                                }
                            }
                        }
                        #endregion

                        query.PageInfo.PageNumber += 1;
                        query.PageInfo.PagingCookie = entCol.PagingCookie;
                        entCol = _servicePrd.RetrieveMultiple(query);
                        count = count + entCol.Entities.Count;
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count of State " + statename + " || " + count);
                    // Console.WriteLine("****************************************** Entity " + entityName + " is ended for State "+ s + ".****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }

            });
        }

        public static void migrateBusMapping(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            string[] stateID =  {
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

            Parallel.ForEach(stateID, state =>
                {
                    string entityName = "hil_businessmapping";
                    //Console.WriteLine("****************************************** Entity " + entityName + " is started for State "+ s + ".****************************************** ");
                    QueryExpression query = new QueryExpression(entityName.ToLower());
                    query.ColumnSet = new ColumnSet("hil_state");
                    query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, state);
                    query.PageInfo = new PagingInfo();
                    query.PageInfo.Count = 5000;
                    query.PageInfo.PageNumber = 1;
                    query.PageInfo.ReturnTotalRecordCount = true;
                    int count = 0;
                    int error = 0;
                    int done = 0;
                    try
                    {

                        EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
                        string statename = entCol[0].GetAttributeValue<EntityReference>("hil_state").Name;
                        count = entCol.Entities.Count;
                        do
                        {
                            #region foreach loop...
                            foreach (Entity entity1 in entCol.Entities)
                            {
                                if (getRecordExist(_serviceDev, entity1.LogicalName, entity1.Id))
                                {
                                    //Console.WriteLine("d");
                                    continue;
                                }
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
                                    _createEntity["createdby"] = systemAdminRef;
                                    if (entity1.Contains("ownerid"))
                                        _createEntity["ownerid"] = systemAdminRef;
                                    _createEntity["modifiedby"] = systemAdminRef;
                                    _createEntity["statecode"] = new OptionSetValue(0);
                                    _createEntity["statuscode"] = new OptionSetValue(1);
                                    string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
                                    if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
                                        _serviceDev.Create(_createEntity);
                                    done++;
                                    Console.WriteLine("Done " + done + "/" + count + " Record of State " + statename);
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
                                        error++;
                                        Console.WriteLine("Error Count " + error);
                                    }
                                }
                            }
                            #endregion

                            query.PageInfo.PageNumber += 1;
                            query.PageInfo.PagingCookie = entCol.PagingCookie;
                            entCol = _servicePrd.RetrieveMultiple(query);
                            count = count + entCol.Entities.Count;
                        }
                        while (entCol.MoreRecords);
                        Console.WriteLine("Count of State " + statename + " || " + count);
                        // Console.WriteLine("****************************************** Entity " + entityName + " is ended for State "+ s + ".****************************************** ");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR !!! " + ex.Message);
                    }

                });
        }
        public static void MigrateData_hil_businessmapping(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
        myGoto:
            string entityName = "hil_businessmapping";
            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(false);
            //query.Criteria.AddCondition("hil_state", ConditionOperator.OnOrAfter, stateId);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            int count = 0;
            int error = 0;
            try
            {
                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
                //Console.WriteLine("Record Count " + entCol.Entities.Count);
                count = entCol.Entities.Count;
                do
                {
                    Console.Write(",");
                    #region foreach loop...
                    //foreach (Entity entity1 in entCol.Entities)
                    //{
                    //    if (getRecordExist(_serviceDev, entity1.LogicalName, entity1.Id))
                    //    {
                    //        //Console.WriteLine("d");
                    //        continue;
                    //    }
                    //    Cont:
                    //    Entity _createEntity = new Entity(entity1.LogicalName, entity1.Id);
                    //    try
                    //    {
                    //        var attributes = entity1.Attributes.Keys;

                    //        foreach (string name in attributes)
                    //        {
                    //            if (name != "modifiedby" && name != "createdby" && name != "organizationid")
                    //            {

                    //                if (entity1[name].GetType().Name == "EntityReference")
                    //                {
                    //                    if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                    //                    {
                    //                        _createEntity[name] = new EntityReference("systemuser", RetriveSystemUser(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name).Id));
                    //                    }
                    //                    else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "uom")
                    //                    {
                    //                        _createEntity[name] = new EntityReference("uom", getBaseUoM(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name).Id));
                    //                    }
                    //                    else
                    //                    {
                    //                        _createEntity[name] = new EntityReference(entity1.GetAttributeValue<EntityReference>(name).LogicalName,
                    //                            CreateRecordIfNotExist(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name), systemAdminRef));
                    //                    }
                    //                }
                    //                else
                    //                {
                    //                    _createEntity[name] = entity1[name];
                    //                }
                    //            }
                    //        }
                    //        _createEntity["createdby"] = systemAdminRef;
                    //        if (entity1.Contains("ownerid"))
                    //            _createEntity["ownerid"] = systemAdminRef;
                    //        _createEntity["modifiedby"] = systemAdminRef;
                    //        _createEntity["statecode"] = new OptionSetValue(0);
                    //        _createEntity["statuscode"] = new OptionSetValue(1);
                    //        string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
                    //        if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
                    //            _serviceDev.Create(_createEntity);
                    //        count++;
                    //        Console.WriteLine("Done " + count + " Record Created On " + entity1["createdon"]);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        if (ex.Message != "Cannot insert duplicate key.")
                    //        {
                    //            Console.WriteLine("Error " + ex.Message);
                    //            if (ex.Message.Contains("Does Not Exist"))
                    //            {
                    //                string[] arr = ex.Message.Split(' ');
                    //                string entityNameRef = arr[1].Replace("'", "");
                    //                string entityIDRef = arr[5];
                    //                migrateSingleRecord(_servicePrd, _serviceDev, entityNameRef, entityIDRef, systemAdminRef);
                    //                goto Cont;
                    //            }
                    //            error++;
                    //            Console.WriteLine("Error Count " + error);
                    //        }
                    //        Console.WriteLine("Error " + ex.Message);
                    //    }
                    //}
                    #endregion
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entCol.PagingCookie;
                    entCol = _servicePrd.RetrieveMultiple(query);
                    count = count + entCol.Entities.Count;
                }
                while (entCol.MoreRecords);
                Console.WriteLine("Count of State " + stateId + " || " + count);
                goto myGoto;
                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
    }
}
