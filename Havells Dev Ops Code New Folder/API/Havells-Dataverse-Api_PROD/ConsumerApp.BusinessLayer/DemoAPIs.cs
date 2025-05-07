using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class DemoAPIs
    {
        public List<Makes> GetMakes()
        {
            List<Makes> jobList = new List<Makes>();
            Makes objMakesOutput;

            try
            {
                CrmServiceClient service = GetCRMConnection();
                if (service != null)
                {
                    string query = String.Format(@"<fetch mapping='logical'>
                    <entity name='{0}'>
                    <attribute name='ogre_name' />
                    <order attribute='createdon' descending='false' />
                    </entity>
                    </fetch>", "ogre_make");
                    EntityCollection images = service.RetrieveMultiple(new FetchExpression(query));

                    foreach (Entity record in images.Entities)
                    {
                        objMakesOutput = new Makes {MakeName= record.GetAttributeValue<string>("ogre_name").ToString(), StatusCode = "200", StatusDescription = "OK" };
                        jobList.Add(objMakesOutput);
                    }
                }
                else
                {
                    objMakesOutput = new Makes { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    jobList.Add(objMakesOutput);
                }
            }
            catch (Exception ex)
            {
                objMakesOutput = new Makes { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                jobList.Add(objMakesOutput);
            }
            return jobList;
        }
        static CrmServiceClient GetCRMConnection()
        {
            CrmServiceClient service = null;
            Guid loginUserGuid = Guid.Empty;

            try
            {
                string authType = "OAuth";
                string userName = "kuldeep.khare@smylsolutions.com";
                string password = "Patanahi@1234";
                string url = "https://org72dd2e85.crm11.dynamics.com/";
                string appId = "51f81489-12ee-4a9e-aaae-a2591f45987d";
                string reDirectURI = "app://58145B91-0C36-4500-8554-080854F2AC97";
                string loginPrompt = "Auto";
                string ConnectionString = string.Format("AuthType = {0};Username = {1};Password = {2}; Url = {3}; AppId={4}; RedirectUri={5};LoginPrompt={6}",
                                                        authType, userName, password, url, appId, reDirectURI, loginPrompt);

                service = new CrmServiceClient(ConnectionString);
                loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
            }
            catch { }
            return service;
        }
    }

    [DataContract]
    public class Makes
    {
        [DataMember]
        public string MakeName { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
