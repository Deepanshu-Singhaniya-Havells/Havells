using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Approvalcreate
{
    internal class Program
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
                //CreateApproval.DeleteApprovals(_service, "06e1df48-690b-ee11-8f6e-6045bdac51bc");//  OCL00029666

                Entity approval = _service.Retrieve("hil_approval", new Guid("0e174bf2-d713-ee11-9cbd-6045bdac55a8"), new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                CreateApproval.SendEmailOnApprovalSubmition(_service,null, approval);
                
                //String EntityName = "hil_orderchecklist";
                //String EntityID = "06e1df48-690b-ee11-8f6e-6045bdac51bc";
                //////String EntityName = "hil_tender";
                //////String EntityID = "c86791b3-5f00-ee11-8f6d-6045bdaf13b7";//tender
                //int purpose = (int)ApprovalPurpose.Order_Price_Approval;
                //int EntobjType = (int)EntityObjectType.Order_Check_List;
                //CreateApproval.CreateApprovals(_service, null, EntityName, EntityID, purpose, EntobjType);

                ////a2d24506-1e10-ee11-9cbd-6045bdac5a1d

                Entity entity = _service.Retrieve("hil_oclproductapproval", new Guid("dde11306-eb10-ee11-9cbd-6045bdac54b1"), new Microsoft.Xrm.Sdk.Query.ColumnSet(true));

                CreateApproval.ActionOnApprovalofOCLPrice(_service, entity, null);
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
