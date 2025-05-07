using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class SAWApprovalAssignToCallCenter
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
                    <entity name='hil_sawactivityapproval'>
                    <attribute name='hil_jobid' />
                    <attribute name='hil_approver' />
                    <order attribute='hil_jobid' descending='false' />
                    <filter type='and'>
                        <condition attribute='hil_approvalstatus' operator='in'>
                        <value>1</value>
                        <value>2</value>
                        <value>5</value>
                        <value>6</value>
                        </condition>
                        <condition attribute='hil_approver' operator='ne' uiname='Vibhor Shukla' uitype='systemuser' value='{C1D89036-5E10-EC11-B6E6-002248D4B299}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_jobid' link-type='inner' alias='a_eecd118653ddea11a813000d3af0563c'>
                        <filter type='and'>
                        <condition attribute='hil_productcategory' operator='ne' uiname='HAVELLS PLUM' uitype='product' value='{B0222990-16FA-E811-A94C-000D3AF06091}' />
                        </filter>
                    </link-entity>
                    <link-entity name='hil_sawactivity' from='hil_sawactivityid' to='hil_sawactivity' link-type='inner' alias='a_c3c5a054cadbea11a813000d3af055b6'>
                        <attribute name='hil_sawcategory' />
                        <filter type='and'>
                        <condition attribute='hil_sawcategory' operator='eq' uiname='Repeat Repair Waiver' uitype='hil_serviceactionwork' value='{B0918D74-44ED-EA11-A815-000D3AF05D7B}' />
                        </filter>
                    </link-entity>
                    </entity>
                </fetch>";
                EntityCollection jobsColl = null;
                Entity _entSAWApproval = null;
                int _rowCount = 1,_totalRowCount=0;
                while (true) {
                    jobsColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (jobsColl.Entities.Count == 0) { break; }
                    _totalRowCount += jobsColl.Entities.Count;
                    foreach (Entity ent in jobsColl.Entities) {
                        _entSAWApproval = new Entity(ent.LogicalName, ent.Id);
                        _entSAWApproval["hil_approver"] = new EntityReference("systemuser",new Guid("C1D89036-5E10-EC11-B6E6-002248D4B299"));
                        _service.Update(_entSAWApproval);
                        Console.WriteLine("Processing SAW Approval..." + _rowCount++.ToString() + "/" + _totalRowCount.ToString());
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
