using System;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class CancelLongAgingJobs
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
                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='msdyn_workorderid' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='msdyn_substatus' operator='not-in'>
                        <value uiname='Canceled' uitype='msdyn_workordersubstatus'>{1527FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='Closed' uitype='msdyn_workordersubstatus'>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='KKG Audit Failed' uitype='msdyn_workordersubstatus'>{6C8F2123-5106-EA11-A811-000D3AF057DD}</value>
                        <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='Work Done SMS' uitype='msdyn_workordersubstatus'>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                      </condition>
                      <condition attribute='createdon' operator='on-or-before' value='2023-09-01' />
                      <condition attribute='hil_brand' operator='eq' value='3' />
                      <condition attribute='hil_callsubtype' operator='eq' uiname='PMS' uitype='hil_callsubtype' value='{E2129D79-3C0B-E911-A94E-000D3AF06CD4}' />
                      <condition attribute='hil_sourceofjob' operator='eq' value='10' />
                    </filter>
                    <link-entity name='hil_address' from='hil_addressid' to='hil_address' visible='false' link-type='outer' alias='a_9f1b877eea04e911a94d000d3af06c56'>
                      <attribute name='hil_state' />
                    </link-entity>
                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='a_7f52360171ebe811a96c000d3af05828'>
                      <attribute name='hil_warrantystatus' />
                      <attribute name='hil_invoicedate' />
                      <attribute name='hil_warrantytilldate' />
                      <attribute name='msdyn_product' />
                    </link-entity>
                  </entity>
                </fetch>";

                EntityCollection userCollection = null;
                Entity entJob = null;
                int i = 1;
                while (true)
                {
                    userCollection = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    foreach (Entity item in userCollection.Entities)
                    {
                        entJob = new Entity(item.LogicalName, item.Id);
                        entJob["hil_webclosureremarks"] = "As per the Service Team, Customer Doesn't want Service.";
                        entJob["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("1527FA6C-FA0F-E911-A94E-000D3AF060A1"));
                        entJob["hil_jobcancelreason"] = new OptionSetValue(7);
                        entJob["hil_closureremarks"] = "As per the Service Team, Customer doesn't want Service.";
                        entJob["hil_cancelticket"] = true;
                        try
                        {
                            _service.Update(entJob);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Processing... " + item.GetAttributeValue<String>("msdyn_name") + " / " + i.ToString() + " /" + ex.Message);
                        }
                        Console.WriteLine("Processing... " +item.GetAttributeValue<String>("msdyn_name") + " / " +  i++.ToString());
                    }
                }
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
