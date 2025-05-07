using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace InventoryTestApp
{
    public class SMLY
    {
        public string connStr = "AuthType=ClientSecret;url={0};ClientId=bc6676d6-1387-4dc4-be89-ba13b08ceb4e;ClientSecret=73P7Q~sWxupzl4j8-B55y5g3QNosxhkjkV6Q2";
        public string orgUrl = "https://orga23838be.crm11.dynamics.com/";
        private IOrganizationService svc = null;
        IOrganizationService service;
        public SMLY()
        {
            service = CreateCrmService();
        }
        public IOrganizationService CreateCrmService()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var conn = new CrmServiceClient(string.Format(connStr, orgUrl));
            return conn.OrganizationWebProxyClient != null ? conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
        }
        public void RunFunction(string[] args)
        {
            upDateVehicleAge();
        }
        public void upDateVehicleAge()
        {
            QueryExpression query = new QueryExpression("ogre_vehicle");
            query.ColumnSet = new ColumnSet("ogre_registrationdate");
            query.Orders.Add(new OrderExpression("createdon", OrderType.Ascending));
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entCol = service.RetrieveMultiple(query);
            int done = 1;
            int count = 0;
            do
            {
                count = count + entCol.Entities.Count;
                foreach (Entity entity in entCol.Entities)
                {
                    try
                    {
                        DateTime registrationDate = entity.GetAttributeValue<DateTime>("ogre_registrationdate");
                        var ss = Math.Round(DateTime.Today.Date.Subtract(registrationDate.Date).Days / (365.2425 / 12), 0);
                        Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                        entity1["ogre_vehicleage"] = ss;
                        service.Update(entity1);
                        Console.WriteLine("Done {0}/{1}", done, count);
                        done++;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entCol.PagingCookie;
                entCol = service.RetrieveMultiple(query);

            }
            while (entCol.MoreRecords);
        }
    }
}
