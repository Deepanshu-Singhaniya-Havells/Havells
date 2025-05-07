using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.DataMigration
{
    public class HelperClass
    {
        public const string SystemAdmin = "5190416c-0782-e911-a959-000d3af06a98";


        protected static Guid RetriveSystemUser(IOrganizationService _servicePrd, IOrganizationService _serviceDev, Guid user)
        {
            Entity prdUser = _servicePrd.Retrieve("systemuser", user, new ColumnSet("domainname"));
            var query = new QueryExpression("systemuser");
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, prdUser.GetAttributeValue<string>("domainname"));
            query.ColumnSet = new ColumnSet(false);
            var results = _serviceDev.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return Guid.Empty;
            else
                return results[0].Id;

        }
        protected static Guid getBaseUoM(IOrganizationService _servicePrd, IOrganizationService _serviceDev, Guid UomId)
        {
            Entity prdUoM = _servicePrd.Retrieve("uom", UomId, new ColumnSet("name"));
            var query = new QueryExpression("uom");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, prdUoM.GetAttributeValue<string>("name"));
            query.ColumnSet = new ColumnSet(false);
            var results = _serviceDev.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return Guid.Empty;
            else
                return results[0].Id;

        }
        protected static void migrateSingleRecord(IOrganizationService _servicePrd, IOrganizationService _serviceDev, string entityName, string ID, EntityReference systemAdminRef)
        {
            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(true);
            query.AddOrder("createdon", OrderType.Ascending);
            if (ID != null)
            {
                query.Criteria.AddCondition(entityName.ToLower() + "id", ConditionOperator.Equal, new Guid(ID));
            }
            else
            {
                query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
            }
            try
            {
                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
                Console.WriteLine("Record Count  entityName " + entityName + " || " + entCol.Entities.Count);
                Parallel.ForEach(entCol.Entities, entity1 =>

                {
                    Guid UomId = Guid.Empty;
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
                                    //else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "product")
                                    //{
                                    //    if (entity1.LogicalName != "product")
                                    //    {
                                    //        _createEntity[name] = new EntityReference(entity1.GetAttributeValue<EntityReference>(name).LogicalName,
                                    //        CreateRecordIfNotExist(_servicePrd, _serviceDev, entity1.GetAttributeValue<EntityReference>(name), systemAdminRef));
                                    //    }
                                    //}
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
                            _serviceDev.Create(_createEntity);
                        Console.WriteLine("Done  entityName " + entityName + " || ");
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message != "Cannot insert duplicate key.")
                        {
                            if (ex.Message.Contains("Does Not Exist"))
                            {
                                string[] arr = ex.Message.Split(' ');
                                string entityNameRef = arr[1].Replace("'", "");
                                string entityIDRef = arr[5];
                                migrateSingleRecord(_servicePrd, _serviceDev, entityNameRef, entityIDRef, systemAdminRef);
                                goto Cont;
                            }

                            Console.WriteLine("Error  entityName " + entityName + " || " + ex.Message);
                        }
                    }
                }
                );
                Console.WriteLine("****************************************** Entity " + entityName + " is ended.****************************************** ");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        protected static Guid CreateRecordIfNotExist(IOrganizationService _servicePrd, IOrganizationService _serviceDev, EntityReference entityReference, EntityReference systemAdminRef)
        {
        Cont1:
            try
            {
                QueryExpression query = new QueryExpression(entityReference.LogicalName);
                query.ColumnSet = new ColumnSet(false);
                query.Criteria.AddCondition(entityReference.LogicalName + "id", ConditionOperator.Equal, entityReference.Id);
                EntityCollection entityCollection = _serviceDev.RetrieveMultiple(query);
                if (entityCollection.Entities.Count != 1)
                {
                    Entity entity = _servicePrd.Retrieve(entityReference.LogicalName, entityReference.Id, new ColumnSet(true));
                    _serviceDev.Create(entity);
                }
            }
            catch (Exception ex)
            {
                if (ex.Message != "Cannot insert duplicate key.")
                {
                    if (ex.Message.Contains("Does Not Exist"))
                    {
                        string[] arr = ex.Message.Split(' ');
                        string entityNameRef = arr[1].Replace("'", ""); ;
                        string entityIDRef = arr[5];
                        if (!ex.Message.Contains("Product"))
                        {
                            migrateSingleRecord(_servicePrd, _serviceDev, entityNameRef, entityIDRef, systemAdminRef);
                            goto Cont1;
                        }
                    }
                    Console.WriteLine("@@@@Error in Entity " + entityReference.LogicalName + " || " + ex.Message);

                }
            }
            return entityReference.Id;
        }
        protected static void MigrateBulkRecords(IOrganizationService _servicePrd, IOrganizationService _serviceDev, String entityName, EntityReference systemAdminRef, bool OnlyActive)
        {
            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(true);
            query.AddOrder("createdon", OrderType.Ascending);
            //query.Criteria.AddCondition("hil_productsubcategory", ConditionOperator.Equal, new Guid("0d1b7022-410b-e911-a94f-000d3af00f43"));
            //if (OnlyActive)
            //    query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            int count = 0;
            int error = 0;
            int recordCount = 0;
            try
            {
                EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
                Console.WriteLine("Record Count entityName " + entityName + " || " + entCol.Entities.Count);
                recordCount = entCol.Entities.Count;
                do
                {
                    #region ForEach...
                    foreach (Entity entity1 in entCol.Entities)
                    {
                        if (getRecordExist(_serviceDev, entity1.LogicalName, entity1.Id))
                        {
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
                            string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
                            if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
                                _serviceDev.Create(_createEntity);
                            Console.WriteLine("Done Count entityName " + entityName + " || " + count);
                            count++;
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message != "Cannot insert duplicate key.")
                            {
                                //Console.WriteLine("Error " + ex.Message);
                                if (ex.Message.Contains("Does Not Exist"))
                                {
                                    string[] arr = ex.Message.Split(' ');
                                    string entityNameRef = arr[1].Replace("'", "");
                                    string entityIDRef = arr[5];
                                    migrateSingleRecord(_servicePrd, _serviceDev, entityNameRef, entityIDRef, systemAdminRef);
                                    goto Cont;
                                }

                                Console.WriteLine("@@@@Error in Entity " + entityName + " || " + ex.Message);
                            }
                        }
                    }
                    #endregion
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entCol.PagingCookie;
                    entCol = _servicePrd.RetrieveMultiple(query);
                    recordCount = recordCount + entCol.Entities.Count;
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
        protected static bool getRecordExist(IOrganizationService _serviceDev, string entityName, Guid ID)
        {
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(true);
            query.AddOrder("createdon", OrderType.Ascending);
            query.Criteria.AddCondition(entityName.ToLower() + "id", ConditionOperator.Equal, ID);
            EntityCollection entCol = _serviceDev.RetrieveMultiple(query);
            if (entCol.Entities.Count == 1)
            {
                return true;
            }
            else
                return false;
        }
        protected static void GetRecordCount(IOrganizationService _servicePrd, IOrganizationService _serviceDev, String entityName)
        {

            Console.WriteLine("Record Count for entityName :- " + entityName);
            int prdCount = 0;
            int QACount = 0;
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.ColumnSet = new ColumnSet(false);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entCol = _servicePrd.RetrieveMultiple(query);
            prdCount = entCol.Entities.Count;
            do
            {
                Console.Write(".");

                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entCol.PagingCookie;
                entCol = _servicePrd.RetrieveMultiple(query);
                prdCount = prdCount + entCol.Entities.Count;

            }
            while (entCol.MoreRecords);
            Console.WriteLine("");
            query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(false);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entColqa = _serviceDev.RetrieveMultiple(query);
            QACount = entColqa.Entities.Count;
            do
            {
                Console.Write(".");

                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entColqa.PagingCookie;
                entColqa = _serviceDev.RetrieveMultiple(query);
                QACount = QACount + entColqa.Entities.Count;

            }
            while (entColqa.MoreRecords);

            Console.WriteLine("");
            Console.WriteLine("Entity " + entityName + " Prd Count !!! " + prdCount);
            Console.WriteLine("Entity " + entityName + " QA Count !!! " + QACount);
        }

        protected static void error(IOrganizationService service)
        {
            try
            {
                string JobProductIDs = "f77e13b4-f40c-ee11-8f6e-6045bdad2773,fd7e13b4-f40c-ee11-8f6e-6045bdad2773,047f13b4-f40c-ee11-8f6e-6045bdad2773";
                string RMAID = "14d2950d-1d22-ee11-9966-6045bdaa91c3";// context.InputParameters["RMAID"].ToString();

                Entity RMAEntity = service.Retrieve("msdyn_rma", new Guid(RMAID), new ColumnSet("ownerid", "msdyn_serviceaccount"));
                EntityReference rmaOwner = RMAEntity.GetAttributeValue<EntityReference>("ownerid");
                Entity account = service.Retrieve(RMAEntity.GetAttributeValue<EntityReference>("msdyn_serviceaccount").LogicalName,
                     RMAEntity.GetAttributeValue<EntityReference>("msdyn_serviceaccount").Id, new ColumnSet("defaultpricelevelid"));
                EntityReference priceList = null;
                if (account.Contains("defaultpricelevelid"))
                    priceList = account.GetAttributeValue<EntityReference>("defaultpricelevelid");
                else
                    throw new InvalidPluginExecutionException("Price List is not mapped.");
                EntityReference wareHouse = GetWarehouse(rmaOwner, service);

                string[] jobproductIdsArray = JobProductIDs.Split(',');
                string value = "";
                foreach (string jobproductId in jobproductIdsArray)
                {
                    value = value + "<value>{" + jobproductId + "}</value>";
                }
                string fetchXML = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                      <entity name=""msdyn_workorderproduct"">
                                        <attribute name=""msdyn_workorderproductid"" />
                                        <attribute name=""hil_replacedpartdescription"" />
                                        <attribute name=""hil_replacedpart"" />
                                        <attribute name=""msdyn_customerasset"" />
                                        <attribute name=""msdyn_quantity"" />
                                        <attribute name=""hil_partamount"" />
                                        <order attribute=""hil_replacedpart"" descending=""false"" />
                                        <filter type=""and"">
                                          <condition attribute=""msdyn_workorderproductid"" operator=""in"">
                                            {value}
                                          </condition>
                                        </filter>
                                        <link-entity name=""product"" from=""productid"" to=""hil_replacedpart"" visible=""false"" link-type=""outer"" alias=""product"">
                                          <attribute name=""description"" />
                                          <attribute name=""defaultuomid"" />
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                foreach (Entity entity in entityCollection.Entities)
                {
                    Entity RMAProduct = new Entity("msdyn_rmaproduct");
                    RMAProduct["msdyn_rma"] = RMAEntity.ToEntityReference();
                    RMAProduct["msdyn_quantitytoreturn"] = entity.GetAttributeValue<double>("msdyn_quantity");
                    RMAProduct["msdyn_woproduct"] = entity.ToEntityReference();
                    RMAProduct["msdyn_product"] = entity.GetAttributeValue<EntityReference>("hil_replacedpart");
                    RMAProduct["msdyn_customerasset"] = entity.GetAttributeValue<EntityReference>("msdyn_cus`1ee`1wwtomerasset");
                    RMAProduct["msdyn_pricelist"] = priceList;
                    RMAProduct["msdyn_processingaction"] = new OptionSetValue(690970001);
                    RMAProduct["ownerid"] = rmaOwner;
                    if (entity.Contains("product.description"))
                        RMAProduct["msdyn_description"] = entity.GetAttributeValue<AliasedValue>("product.description").Value.ToString();
                    if (entity.Contains("product.defaultuomid"))
                        RMAProduct["msdyn_unit"] = (EntityReference)entity.GetAttributeValue<AliasedValue>("product.defaultuomid").Value;
                    RMAProduct["msdyn_returntowarehouse"] = wareHouse;
                    RMAProduct["msdyn_unitamount"] = new Money(entity.GetAttributeValue<Decimal>("hil_partamount"));
                    service.Create(RMAProduct);
                }
                Console.WriteLine("Sucess !");
                Console.WriteLine("RMA lines created");
            }
            catch (Exception ex)
            {
                    Console.WriteLine("Error " + ex.Message);
            }
        }
        static EntityReference GetWarehouse(EntityReference rmaOwner, IOrganizationService service)
        {
            EntityReference wareHouse = null;
            QueryExpression Query = new QueryExpression("bookableresource");
            Query.ColumnSet = new ColumnSet("msdyn_warehouse");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("userid", ConditionOperator.Equal, rmaOwner.Id);
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection bookableColl = service.RetrieveMultiple(Query);
            if (bookableColl.Entities.Count != 1)
            {
                throw new InvalidPluginExecutionException("Bookable resource is not mapped.");
            }
            else
            {
                if (bookableColl[0].Contains("msdyn_warehouse"))
                {
                    wareHouse = bookableColl[0].GetAttributeValue<EntityReference>("msdyn_warehouse");
                }
                else
                {
                    throw new InvalidPluginExecutionException("Warehouse is not mapped.");
                }
            }
            return wareHouse;
        }

    }
}
