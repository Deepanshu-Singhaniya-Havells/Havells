using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


// Below code is used to deled the duplicate records which are created by the multiple hits of the CDR API from the Airtel side.

namespace AE01.Call_Masking
{
    internal class DeleteDuplicateRecords
    {
        private readonly IOrganizationService service; 
        public DeleteDuplicateRecords(IOrganizationService _service)
        {
            this.service = _service; 
        }
        internal void Main()
        {
            EntityCollection tempColl = FetchMultipleRecords(); 
            foreach(var record in tempColl.Entities)
            {   

                string correlationId = record.GetAttributeValue<string>("subject");
                Console.WriteLine("Deleting Record for ID :" + correlationId);
                DeleteDuplicatePhoneCalls(correlationId);
                Console.WriteLine("Deleted duplicate records"); 

            }
        }
        private void DeleteDuplicatePhoneCalls(string CorrelationID)
        {
            List<Entity> duplicateRecords = DetectDuplicatePhoneCalls(CorrelationID);
            Console.WriteLine("Found total records " + duplicateRecords.Count);
            for (int i = 0; i < duplicateRecords.Count - 1; i++)
            {
                service.Delete(duplicateRecords[i].LogicalName, duplicateRecords[i].Id);
            }
        }

        private List<Entity> DetectDuplicatePhoneCalls(string CorrelationID)
        {
            QueryExpression query = new QueryExpression("phonecall");
            query.ColumnSet = new ColumnSet("regardingobjectid");
            query.Criteria.AddCondition("subject", ConditionOperator.Equal, CorrelationID);
            query.Criteria.AddCondition("createdon", ConditionOperator.Today);
            query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        private EntityCollection FetchMultipleRecords()
        {
            QueryExpression query = new QueryExpression("phonecall");
            query.ColumnSet = new ColumnSet("subject", "createdon");
            query.Criteria.AddCondition("subject", ConditionOperator.Like, "Xchange%");
            query.Criteria.AddCondition("createdon", ConditionOperator.Today);
            query.AddOrder("createdon", OrderType.Descending); 
            
            return service.RetrieveMultiple(query); 

        }
    }
}
