using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace D365WebJobs
{
    public class AssignAMCInvoiceTBSH
    {

        internal static void AssignRecord(IOrganizationService service)
        {
            try
            {

                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='10'>
                                  <entity name='invoice'>
                                    <attribute name='name' />
                                    <attribute name='invoiceid' />
                                    <attribute name='hil_modelcode' />
                                    <attribute name='hil_address' />
                                    <attribute name='invoicenumber' />
                                    <attribute name='createdon' />
                                    <order attribute='createdon' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='ownerid' operator='in'>
                                        <value uiname='CRM ADMIN' uitype='systemuser'>{{5190416C-0782-E911-A959-000D3AF06A98}}</value>
                                        <value uiname='Havells India Limited' uitype='systemuser'>{{08074320-FCEE-E811-A949-000D3AF03089}}</value>
                                      </condition>
                                      <condition attribute='hil_address' operator='not-null' />
                                      <condition attribute='hil_modelcode' operator='not-null' />
                                    </filter>
                                  </entity>
                                </fetch>";

                EntityCollection invoiceCollection = null;
                int i = 1;
                do
                {
                    invoiceCollection = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Processing Batch: " + i++ + " Count: " + invoiceCollection.Entities.Count);

                    for(int invoice =  0; invoice < invoiceCollection.Entities.Count; invoice ++)
                    {
        
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Assigning for: " + invoice / invoiceCollection.Entities.Count);

                        string invoiceNumber = invoiceCollection.Entities[invoice].GetAttributeValue<string>("invoicenumber");
                        //Entity invoice = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_address", "hil_modelcode"));
                        if (invoiceCollection.Entities[invoice].Contains("hil_address") && invoiceCollection.Entities[invoice].Contains("hil_modelcode"))
                        {
                            EntityReference addressRef = invoiceCollection.Entities[invoice].GetAttributeValue<EntityReference>("hil_address");
                            Entity address = service.Retrieve(addressRef.LogicalName, addressRef.Id, new ColumnSet("hil_salesoffice"));

                            EntityReference modelRef = invoiceCollection.Entities[invoice].GetAttributeValue<EntityReference>("hil_modelcode");
                            Entity model = service.Retrieve(modelRef.LogicalName, modelRef.Id, new ColumnSet("hil_division"));

                            EntityReference _erDivisionRef = model.GetAttributeValue<EntityReference>("hil_division");
                            EntityReference _erSalesOfficeRef = address.GetAttributeValue<EntityReference>("hil_salesoffice");
                            if (_erDivisionRef != null && _erSalesOfficeRef != null)
                            {
                                QueryExpression query = new QueryExpression("hil_sbubranchmapping");
                                query.ColumnSet = new ColumnSet("hil_branchheaduser");
                                query.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, _erSalesOfficeRef.Id);
                                query.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, _erDivisionRef.Id);
                                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
                                EntityCollection tempColl = service.RetrieveMultiple(query);

                                if (tempColl.Entities.Count > 0)
                                {
                                    EntityReference _erBSH = tempColl.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                                    invoiceCollection.Entities[invoice]["ownerid"] = _erBSH;
                                    service.Update(invoiceCollection.Entities[invoice]);
                                    Console.WriteLine("Assigned Completed for invoice number: " + invoiceNumber); 
                                }
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Divsion or sales Office not present"); 
                            }
                        }
                    }
                    //count = invoiceCollection.Entities.Count;
                }
                while (invoiceCollection.MoreRecords);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red; 
                throw ex;
            }
        }
    }
    
}
