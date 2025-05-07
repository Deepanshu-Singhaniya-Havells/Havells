using System;
using System.Configuration;
using System.Linq;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class BulkRecordDeletion
    {
        public static void BulkDataDeletionJobProduct(IOrganizationService service)
        {
            try
            {
                var _upToDate = ConfigurationManager.AppSettings["UpToDate"].ToString();

                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='200'>
                  <entity name='msdyn_workorderproduct'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_workorderproductid' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_markused' operator='ne' value='1' />
                      <condition attribute='createdon' operator='on-or-before' value='{_upToDate}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ai'>
                      <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='in'>
                          <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                        </condition>
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

                var batchNum = 0;
                var numDeleted = 0;
                while (true)
                {
                    Console.WriteLine("Executing Query..." + DateTime.Now);
                    batchNum += 1;
                    var result = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (result.Entities.Count == 0) {
                        Console.WriteLine("DONE"); break; 
                    }

                    numDeleted += result.Entities.Count;

                    entityRefs = result.Entities.Select(e => e.ToEntityReference());

                    var multiReq = new ExecuteMultipleRequest()
                    {
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = false
                        },
                        Requests = new OrganizationRequestCollection()
                    };

                    var currentList = entityRefs.ToList();

                    currentList.ForEach(r => multiReq.Requests.Add(new DeleteRequest { Target = r }));
                    Console.WriteLine("Deletion Batch Starts..." + DateTime.Now);
                    service.Execute(multiReq);
                    Console.WriteLine("Deletion Batch Ends..." + DateTime.Now);
                    Console.WriteLine("Bulk Record Deletion... " + batchNum + "/" + numDeleted + "/");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }

        public static void BulkDataDeletionJobService(IOrganizationService service)
        {
            try
            {
                var _upToDate = ConfigurationManager.AppSettings["UpToDate"].ToString();

                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='200'>
                  <entity name='msdyn_workorderservice'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_workorderserviceid' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='msdyn_linestatus' operator='ne' value='690970001' />
                      <condition attribute='createdon' operator='on-or-before' value='{_upToDate}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='am'>
                      <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='in'>
                          <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                        </condition>
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

                var batchNum = 0;
                var numDeleted = 0;
                while (true)
                {
                    Console.WriteLine("Executing Query..." + DateTime.Now);
                    batchNum += 1;
                    var result = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (result.Entities.Count == 0)
                    {
                        Console.WriteLine("DONE"); break;
                    }

                    numDeleted += result.Entities.Count;

                    entityRefs = result.Entities.Select(e => e.ToEntityReference());

                    var multiReq = new ExecuteMultipleRequest()
                    {
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = false
                        },
                        Requests = new OrganizationRequestCollection()
                    };

                    var currentList = entityRefs.ToList();

                    currentList.ForEach(r => multiReq.Requests.Add(new DeleteRequest { Target = r }));
                    Console.WriteLine("Deletion Batch Starts..." + DateTime.Now);
                    service.Execute(multiReq);
                    Console.WriteLine("Deletion Batch Ends..." + DateTime.Now);
                    Console.WriteLine("Bulk Record Deletion... " + batchNum + "/" + numDeleted + "/");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }

        public static void BulkDataDeletionJobProductBetweenRange(IOrganizationService service)
        {
            try
            {
                var _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                var _upToDate = ConfigurationManager.AppSettings["UpToDate"].ToString();

                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='200'>
                  <entity name='msdyn_workorderproduct'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_workorderproductid' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_markused' operator='ne' value='1' />
                      <condition attribute='createdon' operator='on-or-after' value='{_fromDate}' />
                      <condition attribute='createdon' operator='on-or-before' value='{_upToDate}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ai'>
                      <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='in'>
                          <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                        </condition>
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

                var batchNum = 0;
                var numDeleted = 0;
                while (true)
                {
                    Console.WriteLine("Executing Query..." + DateTime.Now);
                    batchNum += 1;
                    var result = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (result.Entities.Count == 0)
                    {
                        Console.WriteLine("DONE"); break;
                    }

                    numDeleted += result.Entities.Count;

                    entityRefs = result.Entities.Select(e => e.ToEntityReference());

                    var multiReq = new ExecuteMultipleRequest()
                    {
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = false
                        },
                        Requests = new OrganizationRequestCollection()
                    };

                    var currentList = entityRefs.ToList();

                    currentList.ForEach(r => multiReq.Requests.Add(new DeleteRequest { Target = r }));
                    Console.WriteLine("Deletion Batch Starts..." + DateTime.Now);
                    service.Execute(multiReq);
                    Console.WriteLine("Deletion Batch Ends..." + DateTime.Now);
                    Console.WriteLine("Bulk Record Deletion... " + batchNum + "/" + numDeleted + "/");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }

        public static void BulkDataDeletionJobServiceBetweenRange(IOrganizationService service)
        {
            try
            {
                var _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                var _upToDate = ConfigurationManager.AppSettings["UpToDate"].ToString();

                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='200'>
                  <entity name='msdyn_workorderservice'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_workorderserviceid' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='msdyn_linestatus' operator='ne' value='690970001' />
                      <condition attribute='createdon' operator='on-or-after' value='{_fromDate}' />
                      <condition attribute='createdon' operator='on-or-before' value='{_upToDate}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='am'>
                      <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='in'>
                          <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                        </condition>
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

                var batchNum = 0;
                var numDeleted = 0;
                while (true)
                {
                    Console.WriteLine("Executing Query..." + DateTime.Now);
                    batchNum += 1;
                    var result = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (result.Entities.Count == 0)
                    {
                        Console.WriteLine("DONE"); break;
                    }

                    numDeleted += result.Entities.Count;

                    entityRefs = result.Entities.Select(e => e.ToEntityReference());

                    var multiReq = new ExecuteMultipleRequest()
                    {
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = false
                        },
                        Requests = new OrganizationRequestCollection()
                    };

                    var currentList = entityRefs.ToList();

                    currentList.ForEach(r => multiReq.Requests.Add(new DeleteRequest { Target = r }));
                    Console.WriteLine("Deletion Batch Starts..." + DateTime.Now);
                    service.Execute(multiReq);
                    Console.WriteLine("Deletion Batch Ends..." + DateTime.Now);
                    Console.WriteLine("Bulk Record Deletion... " + batchNum + "/" + numDeleted + "/");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }
    }
}
