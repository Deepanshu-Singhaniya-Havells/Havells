using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class MigrateServiceChannelPartner :HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static MigrateServiceChannelPartner()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void migrateServiceChannelPartner(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            String entityName = "account";
            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("customertypecode", ConditionOperator.In, new object[] {5,6,9 } );
            query.AddOrder("createdon", OrderType.Descending);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            int count = 0;
            int error = 0;
            try
            {
                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
                Console.WriteLine("Record Count " + entCol.Entities.Count);
                do
                {
                    #region foreachloop...
                    foreach (Entity entity1 in entCol.Entities)
                    {
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
                            //_createEntity["statuscode"] = new OptionSetValue(0);
                            //_createEntity["statecode"] = new OptionSetValue(2);
                            _createEntity["createdby"] = systemAdminRef;
                            if (entity1.Contains("ownerid"))
                                _createEntity["ownerid"] = systemAdminRef;
                            _createEntity["modifiedby"] = systemAdminRef;
                            string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
                            if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
                                _serviceDev.Create(_createEntity);
                            Console.WriteLine("Done with id " + _createEntity.Id + " Count " + count);
                            count++;
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
                            }
                            else
                            {
                                _serviceDev.Update(_createEntity);
                            }
                            Console.WriteLine("Error " + ex.Message);
                        }
                    }
                    #endregion
                    Console.Write(".");
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entCol.PagingCookie;
                    entCol = _servicePrd.RetrieveMultiple(query);
                    count = count + entCol.Entities.Count;
                }
                while (entCol.MoreRecords);
                Console.WriteLine("Count !!! " + count);
                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }

    }
}
