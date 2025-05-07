using System;
using System.Linq;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class SFAServiceCallCorrection
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion


        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                EntityCollection entcoll;
                try
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='100'>
                        <entity name='msdyn_customerasset'>
                        <attribute name='msdyn_name' />
                        <attribute name='createdon' />
                        <attribute name='msdyn_customerassetid' />
                        <order attribute='createdon' descending='false' />
                            <filter type='and'>
                            <condition attribute='hil_source' operator='eq' value='9' />
                            <condition attribute='createdon' operator='on-or-after' value='2024-04-01' />
                            <condition attribute='hil_branchheadapprovalstatus' operator='not-in'>
                            <value>1</value>
                            <value>2</value>
                            </condition>
                            <condition attribute='statuscode' operator='ne' value='910590001' />
                            </filter>
                        </entity>
                        </fetch>";

                    int rowCount = 1, totalCount = 0;
                    while (true)
                    {
                        entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        totalCount += entcoll.Entities.Count;
                        if (entcoll.Entities.Count > 0)
                        {
                            foreach (Entity entity in entcoll.Entities)
                            {
                                Console.WriteLine($"Processing... {rowCount++}/{totalCount} Serial Number: {entity.GetAttributeValue<string>("msdyn_name")} Created On: {entity.GetAttributeValue<DateTime>("createdon").ToString()}");
                                Entity entCustomerAsset = new Entity(entity.LogicalName, entity.Id);
                                entCustomerAsset["statuscode"] = new OptionSetValue(910590001);
                                try
                                {
                                    _service.Update(entCustomerAsset);
                                }
                                catch (Exception ex )
                                {
                                    Console.WriteLine("ERROR!!! " + ex.Message);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("Batch Completed...");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
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
    }
}