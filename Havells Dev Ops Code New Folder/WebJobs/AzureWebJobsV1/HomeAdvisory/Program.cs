using System;
using Microsoft.Xrm.Sdk;
using System.IO;
using Microsoft.Xrm.Sdk.Query;
using System.Text;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System.Net;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;

namespace HomeAdvisory
{
    class Program
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
            _service = ConnectToCRM(string.Format(connStr, "https://havellscrmdev1.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady) {
                try
                {
                    string _timeStamp = "19000101000000";
                    #region Daily Sync Jobs
                    GetMasterData.getEnquiryType(_service, _timeStamp);
                    GetMasterData.getProductType(_service, _timeStamp);
                    GetMasterData.getPropertytypeData(_service, _timeStamp);
                    GetMasterData.GetCustomersType(_service, _timeStamp);
                    GetMasterData.GetProductMapping(_service, _timeStamp);
                    GetMasterData.AdvisoryMaster(_service, _timeStamp);
                    //OnDemandJobs.CancellationReasonandConstructionType(_service);
                    #endregion
                    #region D365 To SFA Sync Jobs (5Min)
                    try
                    {
                        D365toSFA.SaveEnquiryDetails(_service);
                        D365toSFA.SaveEnquiryDocument(_service);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    #endregion
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

