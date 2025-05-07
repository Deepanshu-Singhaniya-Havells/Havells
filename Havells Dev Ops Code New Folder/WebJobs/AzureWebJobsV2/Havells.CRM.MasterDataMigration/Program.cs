using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

namespace Havells.CRM.MasterDataMigration
{
    internal class Program
    {
        public static IOrganizationService _SourceService = null;
        public static IOrganizationService _TargetService = null;
        public static EntityReference SystemAdmin = null;// new EntityReference("systemuser", new Guid("1a8fc0e8-7e48-ed11-bba2-6045bdac5a88"));
        public static EntityReference CurrencyRef = null;
        static void Main(string[] args)
        {
            //<add key="connStr" value="AuthType=ClientSecret;url={0};ClientId={1};ClientSecret={2}" />
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();

            var SourceUrl = ConfigurationManager.AppSettings["SourceUrl"].ToString();
            Console.WriteLine("Source Enviroment URL " + SourceUrl);

            var SourceClientID = ConfigurationManager.AppSettings["SourceClientID"].ToString();
            Console.WriteLine("Source Enviroment Client ID " + SourceClientID);

            var SourceClientSecreet = ConfigurationManager.AppSettings["SourceClientSecreet"].ToString();
            Console.WriteLine("Source Enviroment Client Secreet " + SourceClientSecreet);

            var targetUrl = ConfigurationManager.AppSettings["targetUrl"].ToString();
            Console.WriteLine("Target Enviroment URL " + targetUrl);

            var targetClientID = ConfigurationManager.AppSettings["targetClientID"].ToString();
            Console.WriteLine("Target Enviroment Client ID " + targetClientID);

            var TargetClientSecreet = ConfigurationManager.AppSettings["TargetClientSecreet"].ToString();
            Console.WriteLine("Source Enviroment Client Secreet " + TargetClientSecreet);

            var TargetAdminId = ConfigurationManager.AppSettings["TargetAdminId"].ToString();
            Console.WriteLine("Target Enviroment Admin GUID " + TargetAdminId);

            string SourceConnString = string.Format(connStr, SourceUrl, SourceClientID, SourceClientSecreet);
            string TargetConnString = string.Format(connStr, targetUrl, targetClientID, TargetClientSecreet);

            _SourceService = createConnection(SourceConnString);
            _TargetService = createConnection(TargetConnString);

            SystemAdmin = new EntityReference("systemuser", new Guid(TargetAdminId));
            CurrencyRef = getTransactionalCurriency();

            //// 1. Migrate Business Geo Mapping
            //MigrateBussGeoMaping();

            //// 2. Price List
            //PriceList_Migration();

            ////3. SBU Product 
            //SBU_Product_Migration();

            ////3. Claim Type
            //ClaimType_Migration();

            ////4. Claim Category 
            //ClaimCategory_Migration();

            ////5. Division Product
            //Division_Product_Migration();

            ////6. Material Group Product
            //MaterialGroup_Product_Migration();
        }

        static void MigrateBussGeoMaping()
        {
            string entityName = "hil_businessmapping";
            Console.WriteLine("****************************************** Entity " + entityName + " is started.****************************************** ");
            QueryExpression query = new QueryExpression(entityName.ToLower());
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
            //query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, new Guid("27fcb4df-bbf7-e811-a94c-000d3af06091"));

            query.AddOrder("createdon", OrderType.Ascending);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            int count = 0;
            int error = 0;
            try
            {
                EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                Console.WriteLine("Record Count " + entCol.Entities.Count);
                do
                {

                    foreach (Entity entity1 in entCol.Entities)
                    {
                    Cont:
                        Entity _createEntity = new Entity(entity1.LogicalName, entity1.Id);
                        try
                        {
                            var attributes = entity1.Attributes.Keys;

                            foreach (string name in attributes)
                            {
                                if (name != "modifiedby" && name != "createdby")
                                {
                                    if (entity1[name].GetType().Name == "EntityReference")
                                    {
                                        if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                        {
                                            Console.WriteLine(name);
                                            _createEntity[name] = new EntityReference("systemuser", RetriveSystemUser(entity1.GetAttributeValue<EntityReference>(name).Id));
                                        }
                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "uom")
                                        {
                                            _createEntity[name] = new EntityReference("uom", getBaseUoM(entity1.GetAttributeValue<EntityReference>(name).Id));
                                        }
                                        else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "transactioncurrency")
                                        {
                                            _createEntity[name] = CurrencyRef;// new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                        }
                                        else
                                        {
                                            _createEntity[name] = entity1[name];
                                        }
                                    }
                                    else
                                    {
                                        _createEntity[name] = entity1[name];
                                    }
                                }
                            }
                            _createEntity["createdby"] = SystemAdmin;
                            if (entity1.Contains("ownerid"))
                                _createEntity["ownerid"] = SystemAdmin;
                            _createEntity["modifiedby"] = SystemAdmin;
                            string nameUoM = entity1.Contains("name") ? entity1.GetAttributeValue<string>("name") : "";
                            if (nameUoM != "Primary Unit" && nameUoM != "Meter" && nameUoM != "Foot" && nameUoM != "Yard" && nameUoM != "Kilometer" && nameUoM != "Mile" && nameUoM != "Hour" && nameUoM != "Meter" && nameUoM != "Unit")
                                _TargetService.Create(_createEntity);
                            count++;
                            Console.WriteLine("Done " + count + " Record Created On " + entity1["createdon"]);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message != "Cannot insert duplicate key.")
                            {
                                Console.WriteLine("Error " + ex.Message);
                                if (ex.Message.Contains("Does Not Exist"))
                                {
                                    string[] arr = ex.Message.Split(' ');
                                    string entityNameRef = arr[0];
                                    string entityIDRef = arr[4];
                                    migrateData(entityNameRef, entityIDRef);
                                    goto Cont;
                                }
                                error++;
                                Console.WriteLine("Error Count " + error);
                            }

                            Console.WriteLine("Error " + ex.Message);
                        }
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entCol.PagingCookie;
                    entCol = _SourceService.RetrieveMultiple(query);
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
        static void PriceList_Migration()
        {
            int count = 1;
            try
            {
                QueryExpression query = new QueryExpression("pricelevel");
                query.ColumnSet = new ColumnSet(true);
                query.AddOrder("createdon", OrderType.Ascending);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int error = 0;
                try
                {
                    EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                    Console.WriteLine("Record Count " + entCol.Entities.Count);
                    do
                    {
                        foreach (Entity entity in entCol.Entities)
                        {
                            try
                            {
                                var attributes = entity.Attributes.Keys;
                                Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                                foreach (string name in attributes)
                                {
                                    if (name != "modifiedby" && name != "createdby")
                                    {
                                        if (entity[name].GetType().Name == "EntityReference")
                                        {
                                            if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                            {
                                                Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                                if (userid == Guid.Empty)
                                                    userid = SystemAdmin.Id;
                                                _createEntity[name] = new EntityReference("systemuser", userid);
                                            }
                                            else
                                                _createEntity[name] = entity[name];
                                        }
                                        else
                                            _createEntity[name] = entity[name];
                                    }
                                }
                                //_createEntity["statecode"] = new OptionSetValue(2);
                                //_createEntity["statuscode"] = new OptionSetValue(0);
                                _createEntity["transactioncurrencyid"] = CurrencyRef;//new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                //_createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                                //_createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                                _TargetService.Create(_createEntity);
                                count++;
                                Console.WriteLine("Price List With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("ERROR !!! " + ex.Message);
                            }
                        }
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** All Price List Migrated.****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void ClaimType_Migration()
        {
            int count = 1;
            try
            {
                QueryExpression query = new QueryExpression("hil_claimtype");
                query.ColumnSet = new ColumnSet(true);
                query.AddOrder("createdon", OrderType.Ascending);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int error = 0;
                try
                {
                    EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                    Console.WriteLine("Record Count " + entCol.Entities.Count);
                    do
                    {
                        foreach (Entity entity in entCol.Entities)
                        {
                            try
                            {
                                var attributes = entity.Attributes.Keys;
                                Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                                foreach (string name in attributes)
                                {
                                    if (name != "modifiedby" && name != "createdby")
                                    {
                                        if (entity[name].GetType().Name == "EntityReference")
                                        {
                                            if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                            {
                                                Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                                if (userid == Guid.Empty)
                                                    userid = SystemAdmin.Id;
                                                _createEntity[name] = new EntityReference("systemuser", userid);
                                            }
                                            else
                                                _createEntity[name] = entity[name];
                                        }
                                        else
                                            _createEntity[name] = entity[name];
                                    }
                                }
                                //_createEntity["statecode"] = new OptionSetValue(2);
                                //_createEntity["statuscode"] = new OptionSetValue(0);
                                //_createEntity["transactioncurrencyid"] = CurrencyRef;//new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                //_createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                                //_createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                                _TargetService.Create(_createEntity);
                                count++;
                                Console.WriteLine("hil_claimtype With Name " + _createEntity.GetAttributeValue<string>("hil_name") + " is created");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("ERROR !!! " + ex.Message);
                            }
                        }
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** All claim type Migrated.****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void ClaimCategory_Migration()
        {
            int count = 1;
            try
            {
                QueryExpression query = new QueryExpression("hil_claimcategory");
                query.ColumnSet = new ColumnSet(true);
                query.AddOrder("createdon", OrderType.Ascending);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int error = 0;
                try
                {
                    EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                    Console.WriteLine("Record Count " + entCol.Entities.Count);
                    do
                    {
                        foreach (Entity entity in entCol.Entities)
                        {
                            try
                            {
                                var attributes = entity.Attributes.Keys;
                                Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                                foreach (string name in attributes)
                                {
                                    if (name != "modifiedby" && name != "createdby")
                                    {
                                        if (entity[name].GetType().Name == "EntityReference")
                                        {
                                            if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                            {
                                                Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                                if (userid == Guid.Empty)
                                                    userid = SystemAdmin.Id;
                                                _createEntity[name] = new EntityReference("systemuser", userid);
                                            }
                                            else
                                                _createEntity[name] = entity[name];
                                        }
                                        else
                                            _createEntity[name] = entity[name];
                                    }
                                }
                                //_createEntity["statecode"] = new OptionSetValue(2);
                                //_createEntity["statuscode"] = new OptionSetValue(0);
                                //_createEntity["transactioncurrencyid"] = CurrencyRef;//new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                //_createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                                //_createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                                _TargetService.Create(_createEntity);
                                count++;
                                Console.WriteLine("hil_claimcategory With Name " + _createEntity.GetAttributeValue<string>("hil_name") + " is created");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("ERROR !!! " + ex.Message);
                            }
                        }
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** All claim category Migrated.****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void SBU_Product_Migration()
        {
            int count = 1;
            try
            {
                QueryExpression query = new QueryExpression("product");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 1);
                //query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, new Guid("27fcb4df-bbf7-e811-a94c-000d3af06091"));

                query.AddOrder("createdon", OrderType.Ascending);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int error = 0;
                try
                {
                    EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                    Console.WriteLine("Record Count " + entCol.Entities.Count);
                    do
                    {
                        foreach (Entity entity in entCol.Entities)
                        {
                            try
                            {
                                Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
                                var attributes = entity.Attributes.Keys;
                                Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                                foreach (string name in attributes)
                                {
                                    if (name != "modifiedby" && name != "createdby")
                                    {
                                        if (entity[name].GetType().Name == "EntityReference")
                                        {
                                            if (entity.GetAttributeValue<EntityReference>(name).LogicalName != "product")
                                            {
                                                _createEntity[name] = entity[name];
                                            }
                                            else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                            {
                                                Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                                if (userid == Guid.Empty)
                                                    userid = SystemAdmin.Id;
                                                _createEntity[name] = new EntityReference("systemuser", userid);
                                            }
                                            else
                                                _createEntity[name] = entity[name];
                                        }
                                        else
                                            _createEntity[name] = entity[name];
                                    }
                                }
                                _createEntity["statecode"] = new OptionSetValue(2);
                                _createEntity["statuscode"] = new OptionSetValue(0);
                                _createEntity["transactioncurrencyid"] = CurrencyRef;//new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                                _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                                _TargetService.Create(_createEntity);
                                count++;
                                Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                        }
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** Entity is ended.****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void Division_Product_Migration()
        {
            int count = 1;
            try
            {
                QueryExpression query = new QueryExpression("product");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 2);
                //query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, new Guid("27fcb4df-bbf7-e811-a94c-000d3af06091"));

                query.AddOrder("createdon", OrderType.Ascending);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int error = 0;
                try
                {
                    EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                    Console.WriteLine("Record Count " + entCol.Entities.Count);
                    do
                    {
                        foreach (Entity entity in entCol.Entities)
                        {
                            try
                            {
                                Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
                                var attributes = entity.Attributes.Keys;
                                Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                                foreach (string name in attributes)
                                {
                                    if (name != "modifiedby" && name != "createdby")
                                    {
                                        if (entity[name].GetType().Name == "EntityReference")
                                        {
                                            if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "product" && name == "hil_sbu")
                                            {
                                                _createEntity[name] = entity[name];
                                            }
                                            else if (entity.GetAttributeValue<EntityReference>(name).LogicalName != "product")
                                            {
                                                _createEntity[name] = entity[name];
                                            }
                                            else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                            {
                                                Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                                if (userid == Guid.Empty)
                                                    userid = SystemAdmin.Id;
                                                _createEntity[name] = new EntityReference("systemuser", userid);
                                            }
                                            else
                                                _createEntity[name] = entity[name];
                                        }
                                        else
                                            _createEntity[name] = entity[name];
                                    }
                                }
                                _createEntity["statecode"] = new OptionSetValue(2);
                                _createEntity["statuscode"] = new OptionSetValue(0);
                                _createEntity["transactioncurrencyid"] = CurrencyRef;//new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                                _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                                _TargetService.Create(_createEntity);
                                count++;
                                Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
                            }
                            catch (Exception ex)
                            {
                                if (!ex.Message.Contains("The specified product ID conflicts with the product ID "))
                                    Console.WriteLine(ex.Message);
                            }
                        }
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** Entity is ended.****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void MaterialGroup_Product_Migration()
        {
            int count = 1;
            try
            {
                QueryExpression query = new QueryExpression("product");
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3);
                //query.Criteria.AddCondition("hil_state", ConditionOperator.Equal, new Guid("27fcb4df-bbf7-e811-a94c-000d3af06091"));

                query.AddOrder("createdon", OrderType.Ascending);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                int error = 0;
                try
                {
                    EntityCollection entCol = _SourceService.RetrieveMultiple(query);
                    Console.WriteLine("Record Count " + entCol.Entities.Count);
                    do
                    {
                        foreach (Entity entity in entCol.Entities)
                        {
                            try
                            {
                                Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
                                var attributes = entity.Attributes.Keys;
                                Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                                foreach (string name in attributes)
                                {
                                    if (name != "modifiedby" && name != "createdby")
                                    {
                                        if (entity[name].GetType().Name == "EntityReference")
                                        {
                                            if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "product" && (name == "hil_sbu"|| name == "hil_division"))
                                            {
                                                _createEntity[name] = entity[name];
                                            }
                                            else if (entity.GetAttributeValue<EntityReference>(name).LogicalName != "product")
                                            {
                                                _createEntity[name] = entity[name];
                                            }
                                            else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                            {
                                                Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                                if (userid == Guid.Empty)
                                                    userid = SystemAdmin.Id;
                                                _createEntity[name] = new EntityReference("systemuser", userid);
                                            }
                                            else
                                                _createEntity[name] = entity[name];
                                        }
                                        else
                                            _createEntity[name] = entity[name];
                                    }
                                }
                                _createEntity["statecode"] = new OptionSetValue(2);
                                _createEntity["statuscode"] = new OptionSetValue(0);
                                _createEntity["transactioncurrencyid"] = CurrencyRef;//new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                                _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                                _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                                _TargetService.Create(_createEntity);
                                count++;
                                Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
                            }
                            catch (Exception ex)
                            {
                                if (!ex.Message.Contains("The specified product ID conflicts with the product ID "))
                                    Console.WriteLine(ex.Message);
                            }
                        }
                    }
                    while (entCol.MoreRecords);
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** Entity is ended.****************************************** ");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR !!! " + ex.Message);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void migrateData(string entityName, string ID)
        {
            int count = 0;
            int error = 0;

        Cont:
            Entity entity1 = _SourceService.Retrieve(entityName, new Guid(ID), new ColumnSet(true));
            Entity _createEntity = new Entity(entity1.LogicalName, entity1.Id);
            try
            {
                var attributes = entity1.Attributes.Keys;
                Guid UomId = Guid.Empty;
                foreach (string name in attributes)
                {

                    if (name != "adx_partnervisible" && (name != "modifiedby" && name != "createdby" && name != "ownerid" && name != "owninguser" && name != "statecode" && name != "statuscode"))
                    {
                        if (entity1[name].GetType().Name == "EntityReference")
                        {
                            if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                            {
                                Console.WriteLine(name);
                                Guid userid = RetriveSystemUser(entity1.GetAttributeValue<EntityReference>(name).Id);
                                if (userid == Guid.Empty)
                                    userid = SystemAdmin.Id;
                                _createEntity[name] = new EntityReference("systemuser", userid);
                            }
                            else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "uom")
                            {
                                UomId = getBaseUoM(entity1.GetAttributeValue<EntityReference>(name).Id);
                                _createEntity[name] = new EntityReference("uom", UomId);
                            }
                            else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName.ToUpper() == "uomschedule".ToUpper())
                            {
                                if (UomId != Guid.Empty)
                                    _createEntity[name] = new EntityReference("uomschedule", getBaseUoMSchedule(UomId));
                            }
                            else if (entity1.GetAttributeValue<EntityReference>(name).LogicalName == "transactioncurrency")
                            {
                                _createEntity[name] = CurrencyRef;// new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                            }
                            else
                            {
                                _createEntity[name] = entity1[name];
                            }
                        }
                        else
                        {
                            _createEntity[name] = entity1[name];
                        }
                    }
                }
                if (entity1.Contains("ownerid"))
                    _createEntity["ownerid"] = SystemAdmin;
                _TargetService.Create(_createEntity);
                count++;
                Console.WriteLine(entityName + " Record Created On " + entity1["createdon"]);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Cannot insert duplicate key.")
                {
                    Console.WriteLine("Error " + ex.Message);
                    if (ex.Message.Contains("Does Not Exist"))
                    {
                        string[] arr = ex.Message.Split(' ');
                        string entityNameRef = arr[0];
                        string entityIDRef = arr[4];
                        if (!ex.Message.Contains("Product"))
                        {
                            migrateData(entityNameRef, entityIDRef);
                            goto Cont;
                        }
                    }
                    error++;
                    Console.WriteLine("Error Count " + error);
                }

                Console.WriteLine("Error " + ex.Message);
            }



        }
        static EntityReference getTransactionalCurriency()
        {
            var query = new QueryExpression("transactioncurrency");
            query.ColumnSet = new ColumnSet(false);
            var results = _SourceService.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return null;
            else
                return results[0].ToEntityReference();
        }
        static Guid getBaseUoM(Guid UomId)
        {
            Entity prdUoM = _TargetService.Retrieve("uom", UomId, new ColumnSet("name"));
            var query = new QueryExpression("uom");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, prdUoM.GetAttributeValue<string>("name"));
            query.ColumnSet = new ColumnSet(false);
            var results = _SourceService.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return Guid.Empty;
            else
                return results[0].Id;

        }
        static Guid getBaseUoMSchedule(Guid UomId)
        {
            Entity prdUoM = _TargetService.Retrieve("uom", UomId, new ColumnSet("uomscheduleid"));

            if (!prdUoM.Contains("uomscheduleid"))
                return Guid.Empty;
            else
                return prdUoM.GetAttributeValue<EntityReference>("uomscheduleid").Id;

        }
        private static Guid RetriveSystemUser(Guid user)
        {

            Entity prdUser = _TargetService.Retrieve("systemuser", user, new ColumnSet("domainname"));
            var query = new QueryExpression("systemuser");
            query.Criteria.AddCondition("domainname", ConditionOperator.Equal, prdUser.GetAttributeValue<string>("domainname"));
            query.ColumnSet = new ColumnSet(false);
            var results = _SourceService.RetrieveMultiple(query);
            if (results.Entities.Count == 0)
                return Guid.Empty;
            else
                return results[0].Id;

        }
        public static IOrganizationService createConnection(string connectionString)
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
                Console.WriteLine("Connection Created");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;
        }

        static Guid getRecord(string entityLogicalName, Guid productid)
        {
            bool conmt = false;
            try
            {
                QueryExpression query = new QueryExpression(entityLogicalName);
                query.ColumnSet = new ColumnSet(true);
                query.Criteria.AddCondition(entityLogicalName + "id", ConditionOperator.Equal, productid);
                EntityCollection entityCollection = _SourceService.RetrieveMultiple(query);
                if (entityCollection.Entities.Count == 1)
                {
                    return entityCollection[0].Id;
                }
                else
                {
                    Entity entity = _TargetService.Retrieve(entityLogicalName, productid, new ColumnSet(true));


                    var attributes = entity.Attributes.Keys;
                    Entity _createEntity = new Entity(entity.LogicalName, entity.Id);
                    bool conti = false;
                    foreach (string name in attributes)
                    {
                        if (name != "modifiedby" && name != "createdby")
                        {
                            if (entity[name].GetType().Name == "EntityReference")
                            {
                                if (entity.GetAttributeValue<EntityReference>(name).LogicalName != "systemuser")
                                {
                                    //Console.WriteLine(name);
                                    Guid productId = entity.GetAttributeValue<EntityReference>(name).Id;
                                    if (productId != entity.Id)
                                    {
                                        Guid prdId = getRecord(entity.LogicalName, productId);
                                        if (prdId != Guid.Empty)
                                            _createEntity[name] = entity[name];
                                        else
                                            conmt = true;
                                    }
                                }
                                else if (entity.GetAttributeValue<EntityReference>(name).LogicalName == "systemuser")
                                {
                                    Guid userid = RetriveSystemUser(entity.GetAttributeValue<EntityReference>(name).Id);
                                    if (userid == Guid.Empty)
                                        userid = SystemAdmin.Id;
                                    _createEntity[name] = new EntityReference("systemuser", userid);
                                }
                                else
                                    _createEntity[name] = entity[name];
                            }
                            else
                                _createEntity[name] = entity[name];
                        }
                    }
                    if (!conti)
                    {

                        _createEntity["statecode"] = new OptionSetValue(2);
                        _createEntity["statuscode"] = new OptionSetValue(0);
                        _createEntity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("0D5FBED2-D342-ED11-BBA3-000D3AF07BF2"));
                        if (entity.Contains("defaultuomid"))
                        {
                            Guid uomId = getBaseUoM(entity.GetAttributeValue<EntityReference>("defaultuomid").Id);
                            _createEntity["defaultuomscheduleid"] = new EntityReference("uomschedule", getBaseUoMSchedule(uomId));
                            _createEntity["defaultuomid"] = new EntityReference("uom", uomId);
                        }
                        Guid recordid = _TargetService.Create(_createEntity);
                        Console.WriteLine("Product With Name " + _createEntity.GetAttributeValue<string>("name") + " is created");
                        return recordid;
                    }
                    else
                        return Guid.Empty;
                }

            }
            catch (Exception ex)
            {
                Console.Write("___________Error : " + ex.Message);
                return Guid.Empty;
            }
        }
    }
}
