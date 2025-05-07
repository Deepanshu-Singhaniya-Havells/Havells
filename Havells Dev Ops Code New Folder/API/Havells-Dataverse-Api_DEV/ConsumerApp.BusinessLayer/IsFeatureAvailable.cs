using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    public class IsFeatureAvailable
    {
        public static bool ReturnIfFeatureNotAvailable()
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            bool IfAvailable = true;
            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_mobileappclosure");
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Send Boolean_ConsumerAppService");
            Query.AddOrder("createdon", OrderType.Ascending);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                hil_integrationconfiguration enConfig = Found.Entities[0].ToEntity<hil_integrationconfiguration>();
                if (enConfig.hil_mobileappclosure != null)
                {
                    if (enConfig.hil_mobileappclosure.Value == 1)
                    {
                        IfAvailable = true;
                    }
                    else
                    {
                        IfAvailable = false;
                    }
                }
            }
            else
            {
                IfAvailable = false;
            }
            return IfAvailable;
        }

        public static bool ReturnYoutubeConfigs()
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            bool IfAvailable = true;
            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_url", "hil_url2", "hil_username", "hil_description");
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Xamarin_Configuration");
            Query.AddOrder("createdon", OrderType.Ascending);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                hil_integrationconfiguration enConfig = Found.Entities[0].ToEntity<hil_integrationconfiguration>();
                if (enConfig.hil_mobileappclosure != null)
                {
                    if (enConfig.hil_mobileappclosure.Value == 1)
                    {
                        IfAvailable = true;
                    }
                    else
                    {
                        IfAvailable = false;
                    }
                }
            }
            else
            {
                IfAvailable = false;
            }
            return IfAvailable;
        }
    }

    [DataContract]
    public class YoutubeApi
    {
        [DataMember(IsRequired = false)]
        public string PlaylistId { get; set; }
        [DataMember(IsRequired = false)]
        public string PlaylistName { get; set; }
        [DataMember(IsRequired = false)]
        public string YoutubeKey { get; set; }
        [DataMember(IsRequired = false)]
        public string PlaylistImages { get; set; }
        public YoutubeApi ReturnYoutubeConfigs()
        {
            YoutubeApi obj = new YoutubeApi();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            
            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_url", "hil_url2", "hil_username", "hil_description");
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Xamarin_Configuration");
            Query.AddOrder("createdon", OrderType.Ascending);
            EntityCollection Found = service.RetrieveMultiple(Query);
            foreach (hil_integrationconfiguration et in Found.Entities)
            {
                if (et.hil_Url != null)
                {
                    obj.PlaylistId = et.hil_Url;
                }
                if (et.hil_URL2 != null)
                {
                    obj.PlaylistName = et.hil_URL2;
                }
                if (et.hil_Username != null)
                {
                   obj.YoutubeKey = et.hil_Username;
                }
                if (et.hil_description != null)
                {
                    obj.PlaylistImages = et.hil_description;
                }
            }
            return (obj);
        }
    }
}
