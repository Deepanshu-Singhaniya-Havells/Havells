using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualBasic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AE01.Miscellaneous
{

    internal class Report(IOrganizationService _service)
    {
        private readonly IOrganizationService service = _service;

        internal void Caluclate()
        {

            Inventory inventoryObj = new(service);
            inventoryObj.Calcualte();
            Grievance grievanceObj = new(service);
            grievanceObj.Calculate();
        }

    }

    internal class Inventory(IOrganizationService _service)
    {
        private readonly IOrganizationService service = _service;

        internal void Calcualte()
        {
            Entity rma = service.Retrieve("hil_inventoryrma", new Guid("bf78591f-96c2-ef11-b8e8-002248d498e5"), new ColumnSet(true));
            Entity updateRMA = new Entity("hil_inventoryrma", rma.Id);
            updateRMA["statecode"] = new OptionSetValue(1);
            UpdateRequest update = new()
            {
                Target = updateRMA
            };
            update.Parameters.Add("BypassBusinessLogicExecution", "CustomSync,CustomAsync");
            service.Execute(update);    

            QueryExpression query = new QueryExpression("hil_inventorysparebills");
            query.ColumnSet = new ColumnSet("hil_inventorysparebillsid", "hil_name", "createdon", "hil_billdate");
            query.Orders.Add(new OrderExpression("hil_name", OrderType.Ascending));

            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.Conditions.Add(new ConditionExpression("createdon", ConditionOperator.OnOrAfter, new DateTime(2024, 12, 12)));
            query.Criteria.AddFilter(filter);

            LinkEntity link = new LinkEntity("hil_inventorysparebills", "hil_inventorypurchaseorder", "hil_ordernumber", "hil_inventorypurchaseorderid", JoinOperator.LeftOuter);
            link.Columns.AddColumns("hil_approvedon");
            link.EntityAlias = "PO";
            query.LinkEntities.Add(link);

            List<Entity> allRecords = RetrieveAllRecords(query);

            int totalDifference = 0;

            decimal diff = 0; 
            for (int i = 0; i < allRecords.Count; i++)
            {
                if (allRecords[i].Contains("PO.hil_approvedon"))
                {
                    DateTime approvedOn = (DateTime)allRecords[i].GetAttributeValue<AliasedValue>("PO.hil_approvedon").Value;

                    diff += ((allRecords[i].GetAttributeValue<DateTime>("hil_billdate") - approvedOn).Days);

                    diff += (decimal)(((allRecords[i].GetAttributeValue<DateTime>("hil_billdate") - approvedOn).Hours) / 24.00);
                    Console.WriteLine(i + " " + diff);
                }
            }

            Console.WriteLine("Total Difference" + diff);
            Console.WriteLine("TAT: " + (decimal)diff / (decimal)allRecords.Count);
        }

        private List<Entity> RetrieveAllRecords(QueryExpression query)
        {
            List<Entity> allRecords = new List<Entity>();
            string pagingCookie = null;
            int pageNumber = 1;
            int fetchCount = 5000;
            EntityCollection retrievedRecords;

            do
            {
                query.PageInfo.Count = fetchCount;
                query.PageInfo.PageNumber = pageNumber;
                query.PageInfo.PagingCookie = pagingCookie;
                retrievedRecords = service.RetrieveMultiple(query);
                allRecords.AddRange(retrievedRecords.Entities);

                if (retrievedRecords.MoreRecords)
                {
                    pagingCookie = retrievedRecords.PagingCookie;
                    pageNumber++;
                }
            } while (retrievedRecords.MoreRecords);

            return allRecords;
        }
    }


    internal class Grievance
    {
        private readonly IOrganizationService service;
        public Grievance(IOrganizationService _service) => this.service = _service;
        internal void Calculate()
        {
            QueryExpression query = new QueryExpression("incident");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("hil_casedepartment", ConditionOperator.Equal, new Guid("ab3dbc3d-4e6e-ee11-8179-6045bdac526a"));
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, new DateTime(2025, 02, 01));
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, new DateTime(2025, 02, 25));
            EntityCollection tempColl = service.RetrieveMultiple(query);

            decimal sum = 0;

            foreach (var item in tempColl.Entities)
            {
                sum += item.GetAttributeValue<decimal>("hil_tatturnaroundtime");

            }

            Console.WriteLine(tempColl.Entities.Count);
        }
    }
}
