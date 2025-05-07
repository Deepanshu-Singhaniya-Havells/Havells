using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using System.Net;

namespace Havells.Dataverse.Plugins.CommonLibs
{
    public  class DataverseServiceFactory
    {
        private IOrganizationService svc = null;
        IOrganizationService service
        {
            get { return svc ?? (svc = CreateCrmService()); }
        }
        /// <summary>
        /// The method initialize iorganization service from service proxy.
        /// </summary>
        /// <returns></returns>
        public IOrganizationService CreateCrmService()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration("Havells.Dataverse.Plugins.CommonLibs.dll");

            if (config != null)
            {

                string clientId = GetAppSetting(config, "powerapp:ClientId");
                string clientSecret = GetAppSetting(config, "powerapp:ClientSecret");
                string orgUrl = GetAppSetting(config, "powerapp:OrgUrl");
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                var conn = new CrmServiceClient($@"AuthType=ClientSecret;url={orgUrl};ClientId={clientId};ClientSecret={clientSecret}");
                return conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            }
            else {
                return null;
            }
        }

        string GetAppSetting(Configuration config, string key)
        {
            KeyValueConfigurationElement element = config.AppSettings.Settings[key];
            if (element != null)
            {
                string value = element.Value;
                if (!string.IsNullOrEmpty(value))
                    return value;
            }
            return string.Empty;
        }

        public EntityCollection RetrieveMultiple(FetchExpression query)
        {
            return service.RetrieveMultiple(query);
        }
        public Guid Create(Entity e)
        {
            return service.Create(e);
        }
        public void Delete(Entity e)
        {
            service.Delete(e.LogicalName, e.Id);
        }
        public Entity Retrieve(string entityname, Guid entityid, ColumnSet column)
        {

            return service.Retrieve(entityname, entityid, column);
        }
        public EntityCollection RetrieveMultiple(QueryExpression query)
        {
            return service.RetrieveMultiple(query);
        }
        public EntityCollection RetrieveMultiple(IOrganizationService serviceVar, QueryExpression query)
        {
            return serviceVar.RetrieveMultiple(query);
        }
        public EntityCollection RetrieveMultiple(IOrganizationService serviceVar, string query)
        {
            return serviceVar.RetrieveMultiple(new FetchExpression(query));
        }
        /// <summary>
        /// The method executes any crm request and 
        /// send organizationresponse back to calling method.
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        public OrganizationResponse Execute(OrganizationRequest request)
        {
            return service.Execute(request);
        }
        public void Update(Entity e)
        {
            service.Update(e);
        }
    }
}
