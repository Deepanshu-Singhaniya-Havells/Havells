using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using Microsoft.Xrm.Tooling.Connector;
using System.Linq;

namespace Havells.Webjobs.DataWarehouse.DeleteWorkOrders
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
        static string _primaryFieldName = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                WorkOrderArchivedData(_service);
            }
        }
        static void DeleteAllUnusedPartsandService(IOrganizationService service, Guid JobId)
        {
            try
            {
                int _rowCount = 1;
                String fetchPart = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorderproduct'>
                    <attribute name='msdyn_workorderproductid' />
                    <filter type='and'>
                      <condition attribute='msdyn_workorder' operator='eq'  value='{0}' />
                    </filter>
                  </entity>
                </fetch>";
                fetchPart = String.Format(fetchPart, JobId);
                EntityCollection enCollProduct = service.RetrieveMultiple(new FetchExpression(fetchPart));
                foreach (Entity en in enCollProduct.Entities)
                {
                    Console.WriteLine($"Deleting Spare Part Records:{_rowCount++}/{enCollProduct.Entities.Count}");
                    try
                    {
                        service.Delete(en.LogicalName, en.Id);
                    }
                    catch {}
                }

                String fetchService = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorderservice'>
                        <attribute name='msdyn_workorderserviceid' />
                        <filter type='and'>
                          <condition attribute='msdyn_workorder' operator='eq'   value='{0}' />
                        </filter>
                      </entity>
                    </fetch>";
                fetchService = String.Format(fetchService, JobId);
                EntityCollection enCollService = service.RetrieveMultiple(new FetchExpression(fetchService));
                _rowCount=1;
                foreach (Entity en in enCollService.Entities)
                {
                    Console.WriteLine($"Deleting Service Records:{_rowCount++}/{enCollService.Entities.Count}");
                    try
                    {
                        service.Delete(en.LogicalName, en.Id);
                    }
                    catch { }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(" Error in deleting un-unsed parts and services" + ex.Message);
            }
        }
        static void WorkOrderArchivedData(IOrganizationService service)
        {
            try
            {
                var _pageNumber = 1;
                _entityName = ConfigurationManager.AppSettings["hil:entityName"].ToString();
                _fromDate = ConfigurationManager.AppSettings["hil:processDate"].ToString();
                _primaryFieldName = ConfigurationManager.AppSettings["hil:primaryFieldName"].ToString();

                string fetchXml = "";

                var _rowCount = 1;
                var _totalCount = 0;
                string _primaryFieldValue = "";
                string _createdOn = "";
                while (true)
                {
                    //fetchXml = $@"<fetch mapping='logical' page='{_pageNumber}'>
                    fetchXml = $@"<fetch mapping='logical'>
                      <entity name='{_entityName}'>
                        <attribute name='createdon' />
                        <attribute name='{_primaryFieldName}' />
                        <order attribute='createdon' descending='false' />
                        <filter type='and'>
                            <condition attribute='createdon' operator='on' value='{_fromDate}' />
                        </filter>
                      </entity>
                    </fetch>";
                    EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    var entityRefs = result.Entities.Select(e => e.ToEntityReference());
                    if (result.Entities.Count == 0) { break; } else { _totalCount += result.Entities.Count; }

                    foreach (Entity ent in result.Entities)
                    {
                        try
                        {
                            _primaryFieldValue = ent.GetAttributeValue<string>("msdyn_name");
                            _createdOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();

                            DeleteAllUnusedPartsandService(service, ent.Id);

                            service.Delete(ent.LogicalName, ent.Id);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"ERROR:\n{ex.Message} JobId: {_primaryFieldValue} / CreatedOn: {_createdOn}");
                        }
                        Console.WriteLine($"Records:{_rowCount}/{_totalCount} JobId: {_primaryFieldValue} / CreatedOn: {_createdOn}");
                        _rowCount++;
                    }
                    _pageNumber++;
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
