using System;
using System.ServiceModel.Description;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Messages;
using System.Linq;

namespace D365_Data_Archival
{
    class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static string _entityName = string.Empty;
        static string _fromDate = string.Empty;
        static string _toDate = string.Empty;
        static string _topRows = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                //DeleteD365ArchivedData();
                //BulkDeleteD365ArchivedData(_service);
                //BulkDeleteD365JobProductNotUsedData(_service);
                BulkDeleteD365JobServiceNotUsedData(_service);
            }
        }

        static void BulkDeleteD365JobProductNotUsedData(IOrganizationService service)
        {
            try
            {
                _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                _toDate = ConfigurationManager.AppSettings["ToDate"].ToString();
                _topRows = ConfigurationManager.AppSettings["RowCount"].ToString();

                string fetchXml = $@"<fetch mapping='logical' top='{_topRows}'>
                  <entity name='msdyn_workorderproduct'>
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_markused' operator='ne' value='1' />
                      <condition attribute='createdon' operator='on-or-after' value='{_fromDate}' />
                      <condition attribute='createdon' operator='on-or-before' value='{_toDate}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ah'>
                      <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='in'>
                          <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                          <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                          <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                        </condition>
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";

                var result = new EntityCollection();// service.RetrieveMultiple(new FetchExpression(fetchXml));

                var _runCoumt = 1;
                var _rowCount = 1;
                while (true)
                {
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    if (result.Entities.Count == 0) { break; }

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

                    service.Execute(multiReq);

                    Console.WriteLine("Bulk Record Deletion... " + _runCoumt + "/" + _rowCount + "/");
                    _rowCount += currentList.Count;
                    _runCoumt++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }

        static void BulkDeleteD365JobServiceNotUsedData(IOrganizationService service)
        {
            try
            {
                _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                _toDate = ConfigurationManager.AppSettings["ToDate"].ToString();
                _topRows = ConfigurationManager.AppSettings["RowCount"].ToString();

                string fetchXml = $@"<fetch mapping='logical' top='{_topRows}'>
                    <entity name='msdyn_workorderservice'>
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='msdyn_linestatus' operator='ne' value='690970001' />
                        <condition attribute='createdon' operator='on-or-after' value='{_fromDate}' />
                        <condition attribute='createdon' operator='on-or-before' value='{_toDate}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ad'>
                        <filter type='and'>
                        <condition attribute='msdyn_substatus' operator='in'>
                            <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{{1527FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                            <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                            <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{{6C8F2123-5106-EA11-A811-000D3AF057DD}}</value>
                            <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                        </condition>
                        </filter>
                    </link-entity>
                    </entity>
                    </fetch>";

                var result = new EntityCollection();// service.RetrieveMultiple(new FetchExpression(fetchXml));

                var _runCoumt = 1;
                var _rowCount = 1;
                while (true)
                {
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    if (result.Entities.Count == 0) { break; }

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

                    service.Execute(multiReq);

                    Console.WriteLine("Bulk Record Deletion... " + _runCoumt + "/" + _rowCount + "/");
                    _rowCount += currentList.Count;
                    _runCoumt++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }

        static void BulkDeleteD365ArchivedData(IOrganizationService service)
        {
            try
            {

                _entityName = ConfigurationManager.AppSettings["EntityName"].ToString();
                _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                _toDate = ConfigurationManager.AppSettings["ToDate"].ToString();
                _topRows = ConfigurationManager.AppSettings["RowCount"].ToString();

                string fetchXml = $@"<fetch mapping='logical' top='{_topRows}'>
                  <entity name='{_entityName}'>
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='{_fromDate}' />
                        <condition attribute='createdon' operator='on-or-before' value='{_toDate}' />
                    </filter>
                  </entity>
                </fetch>";

                var result = service.RetrieveMultiple(new FetchExpression(fetchXml));

                var _runCoumt = 1;
                var _rowCount = 1;
                while (true)
                {
                    result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    if (result.Entities.Count == 0) { break; }

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

                    service.Execute(multiReq);

                    Console.WriteLine("Bulk Record Deletion... " + _runCoumt + "/" + _rowCount + "/");
                    _rowCount += currentList.Count;
                    _runCoumt++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }
        static void DeleteD365ArchivedData()
        {
            try
            {
                var connectionString = @"AuthType=Office365; Url=https://havells.crm8.dynamics.com/;Username=pwcuser1@havells.com;Password=hqRa6GcMvL5tHw54";
                CrmServiceClient conn = new CrmServiceClient(connectionString);

                IOrganizationService service;
                service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                Console.Write("Enter Entity Schema Name: ");
                _entityName = Console.ReadLine();
                Console.Write("Enter From Date: ");
                _fromDate = Console.ReadLine();
                Console.Write("Enter To Date: ");
                _toDate = Console.ReadLine();

                string query = String.Format(@"<fetch mapping='logical'>
                  <entity name='{0}'>
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='{1}' />
                        <condition attribute='createdon' operator='on-or-before' value='{2}' />
                    </filter>
                  </entity>
                </fetch>", _entityName,_fromDate,_toDate);
                int i = 0, j = 0;
                Console.WriteLine("Record deletion Started.");
                while (true)
                {
                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(query));
                    j += entCol.Entities.Count;
                    if (j == 0) {
                        Console.WriteLine("No Data Found to Delete.");
                        break; 
                    }
                    foreach (Entity record in entCol.Entities)
                    {
                        try
                        {
                            Console.WriteLine(Convert.ToDateTime(record["createdon"].ToString()).AddMinutes(330).ToString() + " :" + i.ToString() + "/" + j.ToString());
                            service.Delete(_entityName, record.Id);
                            i += 1;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error !! " + ex.Message);
                        }
                    }
                }
                Console.WriteLine("Record deletion Ended.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
            }

        }

        static DateTime GetEntityMinDateOfEntry(string _objectType, IOrganizationService service)
        {
            try
            {
                string query = String.Format(@"<fetch top='1'>
                  <entity name='{0}'>
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                  </entity>
                </fetch>", _objectType);

                EntityCollection EntityList = service.RetrieveMultiple(new FetchExpression(query));
                if (EntityList.Entities.Count == 0) { return new DateTime(1900, 1, 1); }
                return EntityList.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330);
            }
            catch (Exception ex)
            {
                return new DateTime(1900, 1, 1);
            }
        }

        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
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
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
