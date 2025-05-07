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
using System.Text;

namespace Havells.Webjobs.MDM.InactiveNotInUserBusinessGeoMapping
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
                InactiveNotInUserBusinessGeoMapping(_service);
            }
        }
        static void InactiveNotInUserBusinessGeoMapping(IOrganizationService service)
        {
            try
            {
                string fetchXml = "";

                var _rowCount = 1;
                var _totalCount = 0;
                while (true)
                {
                    fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_businessmapping'>
                        <attribute name='hil_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_businessmappingid' />
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                          <condition attribute='modifiedon' operator='yesterday' />
                          <condition attribute='statuscode' operator='in'>
                            <value>910590000</value>
                            <value>2</value>
                          </condition>
                          <condition attribute='modifiedby' operator='eq' value='{{0DC1D827-DC64-E911-A96C-000D3AF03089}}' />
                        </filter>
                      </entity>
                    </fetch>";
                    EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (result.Entities.Count == 0) { break; } else { _totalCount += result.Entities.Count; }

                    foreach (Entity ent in result.Entities)
                    {
                        try
                        {
                            Entity _entUpdate = new Entity(ent.LogicalName, ent.Id);
                            _entUpdate["statecode"] = new OptionSetValue(0);
                            _entUpdate["statuscode"] = new OptionSetValue(1);
                            _service.Update(_entUpdate);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"ERROR:\n{ex.Message} BizGeoMapping: {ent.GetAttributeValue<string>("hil_name")}");
                        }
                        Console.WriteLine($"Records:{_rowCount}/{_totalCount} JobId: {ent.GetAttributeValue<string>("hil_name")}");
                        _rowCount++;
                    }
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
