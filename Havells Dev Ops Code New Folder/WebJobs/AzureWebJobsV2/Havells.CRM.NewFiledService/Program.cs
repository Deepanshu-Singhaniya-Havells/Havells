using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.NewFiledService
{
    internal class Program
    {
        static Guid UserId = new Guid("edf84a86-fd81-ed11-81ad-6045bdac5567");
        static EntityReference currency = new EntityReference("transactioncurrency", new Guid("b77689f5-c765-ed11-9564-6045bdad2773"));
        static EntityReference AdminID = null;
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = createConnection(finalString);
            
            WhoAmIRequest request = new WhoAmIRequest();

            WhoAmIResponse response = (WhoAmIResponse)service.Execute(request);
            AdminID = new EntityReference("systemuser", response.UserId);
            CreatePOFromWO(service);
        }
        static void CreatePOFromWO(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("msdyn_workorderproduct");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition(new ConditionExpression("msdyn_workorder", ConditionOperator.Equal, new Guid("7c689f16-db7c-4336-8f33-b4e58c6e6e53")));
            EntityCollection entityCOll = service.RetrieveMultiple(query);
            foreach (Entity entity in entityCOll.Entities)
            {
                Entity entityPO = new Entity("msdyn_purchaseorder");
                entityPO["msdyn_workorder"] = entity.GetAttributeValue<EntityReference>("msdyn_workorder");
                entityPO["msdyn_vendor"] = getVendorAccount(service, UserId);
                entityPO["hil_potype"] = new OptionSetValue(1);
                entityPO["hil_approver"] = getApprover(service,UserId);
                Entity BookableResourceEntity = getBookableResource(service, UserId);
                entityPO["msdyn_requestedbyresource"] = BookableResourceEntity.ToEntityReference();
                entityPO["msdyn_receivetowarehouse"] = BookableResourceEntity.GetAttributeValue<EntityReference>("msdyn_warehouse");
                entityPO["transactioncurrencyid"] = currency;
                entityPO["msdyn_approvalstatus"] = new OptionSetValue(690970000);
                entityPO["msdyn_approvedrejectedby"] = AdminID;
                Guid POID = service.Create(entityPO);

            }
        }
        static EntityReference getVendorAccount(IOrganizationService service, Guid User)
        {
            EntityReference account = null;
            QueryExpression query = new QueryExpression("account");
            query.ColumnSet = new ColumnSet("accountid");
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, User);
            EntityCollection accountColl = service.RetrieveMultiple(query);
            if (accountColl.Entities.Count >= 1)
            {
                account = accountColl[0].ToEntityReference();
            }
            return account;
        }
        static EntityReference getApprover(IOrganizationService service, Guid User)
        {
            Entity userEntity= service.Retrieve("systemuser",User,new ColumnSet("parentsystemuserid"));
            return userEntity.GetAttributeValue<EntityReference>("parentsystemuserid");
        }
        static Entity getBookableResource(IOrganizationService service, Guid User)
        {
            Entity BookableResourceEntity = null;
            QueryExpression query = new QueryExpression("bookableresource");
            query.ColumnSet = new ColumnSet("msdyn_warehouse");
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("userid", ConditionOperator.Equal, User);
            EntityCollection accountColl = service.RetrieveMultiple(query);
            if (accountColl.Entities.Count >= 1)
            {
                BookableResourceEntity = accountColl[0];
            }
            return BookableResourceEntity;
        }
        static IOrganizationService createConnection(string connectionString)
        {
            IOrganizationService organizationService = null;
            try
            {
                organizationService = (IOrganizationService)new CrmServiceClient(connectionString);
                if (((CrmServiceClient)organizationService).LastCrmException != null && (((CrmServiceClient)organizationService).LastCrmException.Message == "OrganizationWebProxyClient is null" || ((CrmServiceClient)organizationService).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)organizationService).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)organizationService).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }

            return organizationService;
        }
    }
}
