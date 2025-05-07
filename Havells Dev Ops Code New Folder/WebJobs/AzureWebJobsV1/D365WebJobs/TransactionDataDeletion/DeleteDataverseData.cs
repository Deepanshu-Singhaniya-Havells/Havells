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

namespace D365WebJobs
{
    public class DeleteDataverseData
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static string _entityName = string.Empty;
        static string _fromDate = string.Empty;
        static string _toDate = string.Empty;

        #endregion

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                BulkDeleteD365WorkOrderManagementData(_service);
            }
        }
        static void BulkDeleteD365WorkOrderManagementData(IOrganizationService service)
        {
            try
            {
                var _entityName = ConfigurationManager.AppSettings["EntityName"].ToString();
                var _dateAfter = ConfigurationManager.AppSettings["DateAfter"].ToString();
                var _dateBefore = ConfigurationManager.AppSettings["DateBefore"].ToString();

                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='400'>
                  <entity name='{_entityName}'>
                    <attribute name='createdon' />
                    <attribute name='msdyn_workorderproductid' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_markused' operator='ne' value='1' />
                      <condition attribute='createdon' operator='on-or-after' value='{_dateAfter}' />
                      <condition attribute='createdon' operator='on-or-before' value='{_dateBefore}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ad'>
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

                int batchNum = 0;
                int numDeleted = 0;
                while (true)
                {
                    batchNum += 1;
                    Console.WriteLine("Query Execution Starts... " + DateTime.Now.ToString());
                    var result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    Console.WriteLine("Query Execution Ends... " + DateTime.Now.ToString());
                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());

                    if (result.Entities.Count == 0) { break; }

                    numDeleted += result.Entities.Count;

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

                    
                    Console.WriteLine("Bulk Record Deletion... " + batchNum + "/" + numDeleted + "/");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error !! " + ex.Message);
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
