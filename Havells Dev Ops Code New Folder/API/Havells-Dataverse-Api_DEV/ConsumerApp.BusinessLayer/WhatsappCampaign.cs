using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class WhatsappCampaign
    {
        public KKGResponse UpdateKKGConsentOnJob(KKGRequest kKGRequest)
        {
            KKGResponse kKGResponse = new KKGResponse();
            try
            {
                if (string.IsNullOrWhiteSpace(kKGRequest.JobId))
                {
                    kKGResponse.status_code = 204;
                    kKGResponse.status_description = "Job Id is required.";
                    return kKGResponse;
                }
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (((CrmServiceClient)service).IsReady)
                    {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_jobsextension'>
                                            <attribute name='hil_jobsextensionid'/>
                                            <attribute name='hil_name'/>
                                            <attribute name='createdon'/>
                                            <attribute name='hil_jobs'/>
                                            <order attribute='hil_name' descending='false'/>
                                            <filter type='and'>
                                                <condition attribute='hil_name' operator='eq' value='{kKGRequest.JobId}'/>
                                            </filter>
                                            </entity>
                                            </fetch>";
                        EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            Entity jobsextension = entcoll.Entities[0];
                            jobsextension["hil_kkgconsent"] = kKGRequest.UserResponse;
                            service.Update(jobsextension);
                            kKGResponse.status_code = 200;
                            kKGResponse.status_description = "Success";
                        }
                        else
                        {
                            kKGResponse.status_code = 204;
                            kKGResponse.status_description = "Job Id does not exist.";
                        }
                    }
                }
                else
                {
                    kKGResponse.status_code = 503;
                    kKGResponse.status_description = "D365 Service Unavailable.";
                }
            }
            catch (Exception ex)
            {
                kKGResponse.status_code = 500;
                kKGResponse.status_description = "Internal Server Error : " + ex.Message.ToUpper();
            }
            return kKGResponse;
        }
    }

    [DataContract]
    public class KKGRequest
    {
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public bool UserResponse { get; set; }
    }
    [DataContract]
    public class KKGResponse
    {
        [DataMember]
        public int status_code { get; set; }
        [DataMember]
        public string status_description { get; set; }
    }

}
