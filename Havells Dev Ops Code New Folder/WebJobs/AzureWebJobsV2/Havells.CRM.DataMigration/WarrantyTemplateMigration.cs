using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class WarrantyTemplateMigration : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static WarrantyTemplateMigration()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void warrantyTemplateMigration(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            string[] entityList = {
                "campaign"
                //"hil_warrantytemplate",
               // "hil_servicebom",
                //"hil_productcatalog",
                //"hil_observation",
                //"hil_warrantyscheme",
                //"hil_schemeline",
                //"hil_pmsconfiguration",
                //"hil_labor",
                //"hil_part"
                //"msdyn_incidenttype",
                //"msdyn_incidenttypeproduct",
                //"msdyn_incidenttypeservice"
            };
            Parallel.ForEach(entityList, entityName =>
            {
                MigrateBulkRecords(_servicePrd, _serviceDev, entityName, systemAdminRef, true);
            });
        }
        public static void MoveIncdentTypeProductAndService(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            QueryExpression query = new QueryExpression("msdyn_incidenttype");
            query.ColumnSet = new ColumnSet(true);
            query.AddOrder("createdon", OrderType.Ascending);
            query.Criteria.AddCondition("hil_model", ConditionOperator.Equal, new Guid("0d1b7022-410b-e911-a94f-000d3af00f43"));
            query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
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
                Console.WriteLine("Record Count entityName  || " + entCol.Entities.Count);
                recordCount = entCol.Entities.Count;
                do
                {
                    foreach (Entity entity1 in entCol.Entities)
                    {
                        string[] entityList = {
                            "msdyn_incidenttypeproduct",
                            "msdyn_incidenttypeservice"
                        };
                        foreach (string entityName in entityList)
                        {
                            QueryExpression query1 = new QueryExpression(entityName);
                            query1.ColumnSet = new ColumnSet(true);
                            query1.AddOrder("createdon", OrderType.Ascending);
                            query1.Criteria.AddCondition("msdyn_incidenttype", ConditionOperator.Equal, entity1.Id);
                            query1.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                            query1.PageInfo = new PagingInfo();
                            query1.PageInfo.Count = 5000;
                            query1.PageInfo.PageNumber = 1;
                            query1.PageInfo.ReturnTotalRecordCount = true;
                            EntityCollection msdyn_incidenttypeserviceColl = _servicePrd.RetrieveMultiple(query1);
                            Console.WriteLine("Record Count entityName  || " + msdyn_incidenttypeserviceColl.Entities.Count);
                            int msdyn_incidenttypeserviceRecordCount = msdyn_incidenttypeserviceColl.Entities.Count;
                            do
                            {
                                foreach (Entity entity12 in msdyn_incidenttypeserviceColl.Entities)
                                {
                                    CreateRecordIfNotExist(_servicePrd, _serviceDev, entity12.ToEntityReference(), systemAdminRef);
                                }
                                query1.PageInfo.PageNumber += 1;
                                query1.PageInfo.PagingCookie = msdyn_incidenttypeserviceColl.PagingCookie;
                                msdyn_incidenttypeserviceColl = _servicePrd.RetrieveMultiple(query1);
                                msdyn_incidenttypeserviceRecordCount = msdyn_incidenttypeserviceRecordCount + msdyn_incidenttypeserviceColl.Entities.Count;
                            }
                            while (msdyn_incidenttypeserviceColl.MoreRecords);
                        }
                        while (entCol.MoreRecords) ;
                    }
                    Console.WriteLine("Count !!! " + count);
                    Console.WriteLine("****************************************** Entity  is ended.****************************************** ");
                }
                while (entCol.MoreRecords);
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entCol.PagingCookie;
                entCol = _servicePrd.RetrieveMultiple(query);
                recordCount = recordCount + entCol.Entities.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
    }
}
