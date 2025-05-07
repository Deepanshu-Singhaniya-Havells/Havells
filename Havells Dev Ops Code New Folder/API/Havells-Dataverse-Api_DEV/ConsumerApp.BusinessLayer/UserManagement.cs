using System;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.Xrm.Sdk.Client;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class UserManagement
    {
        [DataMember]
        public Guid UserGuid { get; set; }
        [DataMember]
        public Guid BusinessGuid { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        private Uri GetServiceUrl(string organizationUrl)
        {
            return new Uri(organizationUrl + @"/xrmservices/2011/organization.svc/web?SdkClientVersion=8.2");
        }
        public UserManagement UpdateBusinessUnit(UserManagement userData)
        {
            UserManagement objUserManagement = null;
            //try
            //{
            //    IOrganizationService service = ConnectToCRM.GetOrgServiceHILTest();
            //    if (service != null)
            //    {
            //        if (userData.UserGuid == Guid.Empty)
            //        {
            //            objUserManagement = new UserManagement { StatusCode = "204", StatusDescription = "User Guid is required." };
            //            return objUserManagement;
            //        }
            //        if (userData.BusinessGuid == Guid.Empty)
            //        {
            //            objUserManagement = new UserManagement { StatusCode = "204", StatusDescription = "Business Unit is required." };
            //            return objUserManagement;
            //        }
            //        try
            //        {
            //            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            //            {
            //                SetBusinessSystemUserRequest changeUserBURequest = new SetBusinessSystemUserRequest();
            //                changeUserBURequest.BusinessId = userData.BusinessGuid;
            //                changeUserBURequest.UserId = userData.UserGuid;
            //                changeUserBURequest.ReassignPrincipal = new EntityReference(SystemUser.EntityLogicalName, userData.UserGuid);
            //                service.Execute(changeUserBURequest);
            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //            return new UserManagement { UserGuid = userData.UserGuid, BusinessGuid = userData.BusinessGuid, StatusCode = "204", StatusDescription = ex.Message };
            //        }
            //        return new UserManagement { UserGuid = userData.UserGuid, BusinessGuid = userData.BusinessGuid, StatusCode = "200", StatusDescription = "OK" };
            //    }
            //    else
            //    {
            //        return new UserManagement { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
            //    }
            //}
            //catch (Exception ex)
            //{
            //    return new UserManagement { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            //}

            try
            {
                if (userData.UserGuid == Guid.Empty)
                {
                    objUserManagement = new UserManagement { StatusCode = "204", StatusDescription = "User Guid is required." };
                    return objUserManagement;
                }
                if (userData.BusinessGuid == Guid.Empty)
                {
                    objUserManagement = new UserManagement { StatusCode = "204", StatusDescription = "Business Unit is required." };
                    return objUserManagement;
                }

                string organizationUrl = "https://greenlightcorp1.crm.dynamics.com/";
                string clientId = "d4553866-2454-4095-8090-4fe91201fa9e";
                string appKey = "876V.4FFKdGfz35Qx6~LZ1p25~u16~63zD";
                string aadInstance = "https://login.microsoftonline.com/";
                string tenantID = "57cdd3cf-ae8b-4672-aac8-72237a79b841";

                ClientCredential clientcred = new ClientCredential(clientId, appKey);
                AuthenticationContext authenticationContext = new AuthenticationContext(aadInstance + tenantID);
                AuthenticationResult authenticationResult = authenticationContext.AcquireTokenAsync(organizationUrl, clientcred).Result;
                var requestedToken = authenticationResult.AccessToken;

                try
                {
                    using (var sdkService = new OrganizationWebProxyClient(GetServiceUrl(organizationUrl), false))
                    {
                        sdkService.HeaderToken = requestedToken;

                        OrganizationRequest request = new OrganizationRequest()
                        {
                            RequestName = "WhoAmI"
                        };
                        WhoAmIResponse response = sdkService.Execute(new WhoAmIRequest()) as WhoAmIResponse;

                        SetBusinessSystemUserRequest changeUserBURequest1 = new SetBusinessSystemUserRequest();
                        changeUserBURequest1.BusinessId = userData.BusinessGuid;
                        changeUserBURequest1.UserId = userData.UserGuid;
                        changeUserBURequest1.ReassignPrincipal = new EntityReference(SystemUser.EntityLogicalName, userData.UserGuid);
                        sdkService.Execute(changeUserBURequest1);
                    }
                }
                catch (Exception ex) {
                    objUserManagement = new UserManagement
                    {
                        StatusCode = "500",
                        StatusDescription = ex.Message.ToUpper()
                    };
                }
                objUserManagement = new UserManagement { UserGuid = userData.UserGuid, BusinessGuid = userData.BusinessGuid, StatusCode = "200", StatusDescription = "OK" };
            }
            catch (Exception ex)
            {
                objUserManagement = new UserManagement
                {
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper()
                };
            }
            return objUserManagement;
        }
    }
}
