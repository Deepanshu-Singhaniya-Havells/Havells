using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using HavellsNewPlugin.SAP_IntegrationForOrderCreation;
using System.Text;
using RestSharp;
using Newtonsoft.Json;

namespace HavellsNewPlugin.TenderModule.DOA_Approval
{
    public class CustomerOutstanding : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                decimal creditLimit = 0;
                int creditDays = 0;
                decimal OutstandingAmount = 0;
                decimal advance = 0;
                decimal overDueAmount = 0;
                tracingService.Trace("Send Approval Email plugin started.");
                var entityName = context.InputParameters["EntityName"].ToString();
                var entityId = context.InputParameters["EntityID"].ToString();
                tracingService.Trace("entityName      " + entityName);

                Entity entity = service.Retrieve(entityName, new Guid(entityId), new ColumnSet("hil_customercode"));
                String customerCode = entity.GetAttributeValue<string>("hil_customercode");

                IntegrationConfig intConfigCreditCheck = Models.IntegrationConfiguration(service, "CreditCheck");

                RootCreditCheck requestCreditCheck = new RootCreditCheck();
                requestCreditCheck.I_KKBER = "100";
                requestCreditCheck.I_KUNNR = customerCode;
                requestCreditCheck.I_REGUL = "HIL";
                requestCreditCheck.I_VKORG = "X";

                decimal overAllValue = 0;

                QueryExpression query = new QueryExpression("hil_oaproduct");
                query.ColumnSet = new ColumnSet("hil_netvalue");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_oaheader", ConditionOperator.Equal, entity.Id));

                EntityCollection oaPrd = service.RetrieveMultiple(query);
                foreach (Entity oaprd in oaPrd.Entities)
                {
                    overAllValue = overAllValue + oaprd.GetAttributeValue<Money>("hil_netvalue").Value;
                }


                
                string dataCreditCheck = JsonConvert.SerializeObject(requestCreditCheck);

                RestResponse resposeCreditCheck = Integration(intConfigCreditCheck, dataCreditCheck);
               // throw new InvalidPluginExecutionException("resposeCreditCheck " + resposeCreditCheck.Content);
                RespnseCreditCheck returndatas = JsonConvert.DeserializeObject<RespnseCreditCheck>(resposeCreditCheck.Content);
                creditLimit = decimal.Parse(returndatas.CRD_LIMIT_AMT.Trim());
                creditDays = int.Parse(returndatas.CREDIT_DAYS.Trim());

                IntegrationConfig intConfigGetOpenItems = Models.IntegrationConfiguration(service, "GetOpenItems");

                RootGetOpenItems requestGetOpenItems = new RootGetOpenItems();
                DateTime currentDate = DateTime.Now.Date;

                String timestamp = currentDate.Year.ToString() + "-" + currentDate.Month.ToString().PadLeft(2, '0') + "-" + currentDate.Day.ToString().PadLeft(2, '0');
                tracingService.Trace("timestamp  " + timestamp);


                requestGetOpenItems.KEYDATE = timestamp;
                requestGetOpenItems.CUSTOMER = customerCode;

                string dataGetOpenItems = JsonConvert.SerializeObject(requestGetOpenItems);

                RestResponse resposeGetOpenItems = Integration(intConfigGetOpenItems, dataGetOpenItems);

                RespnseGetOpenItems returnGetOpenItems = JsonConvert.DeserializeObject<RespnseGetOpenItems>(resposeGetOpenItems.Content);
                if (returnGetOpenItems.NET_DUE != "")
                {
                    if (returnGetOpenItems.NET_DUE.Trim().Contains("-"))
                    {
                        string amount = returnGetOpenItems.NET_DUE.Trim().Replace("-", "");
                        advance = decimal.Parse(amount);
                    }
                    else
                        OutstandingAmount = decimal.Parse(returnGetOpenItems.NET_DUE.Trim());
                }

                if (returnGetOpenItems.OVERDUE != "")
                    overDueAmount = decimal.Parse(returnGetOpenItems.OVERDUE.Trim());


                Entity oaHeader = new Entity(entity.LogicalName, entity.Id);
                oaHeader["hil_tolerance"] = decimal.Parse("5.00");
                oaHeader["hil_ordervalue"] = new Money(overAllValue);
                oaHeader["hil_advanceamount"] = new Money(advance);
                decimal tolaranceAmount = (5 * overAllValue / 100);
                decimal orderValue = overAllValue + tolaranceAmount;
                decimal orderValueGST = (18 * orderValue / 100);

                decimal limit = orderValueGST + orderValue;
                //throw new InvalidPluginExecutionException("overAllValue: "+ overAllValue+"  tolaranceAmount " + tolaranceAmount.ToString() + " tolaranceGST: " + tolaranceGST.ToString() + " limit: " + limit);
                oaHeader["hil_limitrequested"] = new Money(limit);
                oaHeader["hil_creditlimit"] = new Money(creditLimit);
                oaHeader["hil_creditdays"] = creditDays;
                oaHeader["hil_outstandingamount"] = new Money(OutstandingAmount);
                service.Update(oaHeader);

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
        public RestResponse Integration(IntegrationConfig intConfig, string data)
        {
            string _authInfo = intConfig.Auth;
            _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
            _authInfo = "Basic " + _authInfo;

            var client = new RestClient(intConfig.uri);
            //client.Timeout = -1;
            var request = new RestRequest();
            request.AddHeader("Authorization", _authInfo);
            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("application/json", data, ParameterType.RequestBody);
            RestResponse response = client.Execute(request, Method.Post);
            return response;
        }
    }
}
