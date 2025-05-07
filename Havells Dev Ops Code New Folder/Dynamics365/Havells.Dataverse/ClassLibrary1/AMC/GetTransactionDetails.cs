using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;


namespace Havells.Dataverse.CustomConnector.AMC
{
    public class GetTransactionDetails : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]);
            string UserToken = Convert.ToString(context.InputParameters["UserToken"]);

            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            var data = JsonSerializer.Deserialize<AMCOrdersParam>(jsonString);
            string SourceType = data.SourceType;
            string CustomerGuid = data.CustomerGuId;

            if (!APValidate.IsvalidGuid(CustomerGuid))
            {
                string msg = string.IsNullOrWhiteSpace(CustomerGuid) ? "Customer Guid required." : "Please enter valid Customer Guid.";
                var responnse = JsonSerializer.Serialize(new { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (string.IsNullOrWhiteSpace(SourceType))
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Source type is required." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            else if (SourceType != "6")
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Please enter valid Source Type." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            else
            {
                AMCOrdersParam objparam = new AMCOrdersParam();
                objparam.CustomerGuId = CustomerGuid;
                objparam.SourceType = SourceType;
                var response = GetTransactionDetail(service, objparam, LoginUserId);
                if (response.Item2.StatusCode == (int)HttpStatusCode.OK)
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new { StatusCode = 200, response.Item1 }).Replace("\"Item1\":", "");
                else
                    context.OutputParameters["data"] = JsonSerializer.Serialize(response.Item2);
                return;
            }
        }
        public (List<TranscationHistory>, RequestStatus) GetTransactionDetail(IOrganizationService _crmService, AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            List<TranscationHistory> lstTranscationHistory = new List<TranscationHistory>();
            try
            {
                if (_crmService != null)
                {
                    try
                    {
                        new Guid(AMCOrdersParam.CustomerGuId);
                    }
                    catch (Exception)
                    {
                        return (lstTranscationHistory, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid customer Guid."
                        });
                    }
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='salesorder'>
                                        <attribute name='salesorderid'/>
                                        <attribute name='name'/>
                                        <attribute name='customerid'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='statuscode'/>
                                        <attribute name='hil_paymentstatus'/>
                                        <attribute name='hil_productdivision'/>
                                        <attribute name='ownerid'/>
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                            <condition attribute='customerid' operator='eq'  value='{AMCOrdersParam.CustomerGuId}'/>                                           
                                        </filter>
                                        </entity>
                                        </fetch>";
                    EntityCollection entityColl = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entityColl.Entities.Count > 0)
                    {
                        foreach (Entity ent in entityColl.Entities)
                        {
                            TranscationHistory transcationHistory = new TranscationHistory();
                            transcationHistory.InvoiceId = ent.GetAttributeValue<string>("name");
                            transcationHistory.PlanName = ent.Contains("hil_productdivision") ? ent.GetAttributeValue<EntityReference>("hil_productdivision").Name : "";
                            transcationHistory.Amount = Math.Round(ent.Contains("totalamount") ? ent.GetAttributeValue<Money>("totalamount").Value : 0, 2).ToString();

                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                                                  <entity name='hil_paymentreceipt'>
                                                    <attribute name='hil_paymentreceiptid' />
                                                    <attribute name='hil_transactionid' />
                                                    <attribute name='hil_paymentstatus' />
                                                    <attribute name='hil_bankreferenceid' />                                                    
                                                    <attribute name='createdon' />
                                                    <order attribute='createdon' descending='true' />
                                                    <filter type='and'>
                                                        <condition attribute='statecode' operator='eq' value='0' />
                                                        <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";
                            EntityCollection entPaymentreceipt = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entPaymentreceipt.Entities.Count > 0)
                            {
                                int Status = entPaymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value;
                                if (Status == 1 || Status == 3)
                                {
                                    transcationHistory.PaymentStatus = "Pending";
                                    transcationHistory.InfoMessage = "InCase of failure the amount will be deposited back to your account in 24 hours";
                                }
                                if (Status == 2 || Status == 5)
                                {
                                    transcationHistory.PaymentStatus = "Failed";
                                    transcationHistory.InfoMessage = "Refund initiated: the amount will be reflected in your account in 24 hours";
                                }
                                if (Status == 4)
                                {
                                    transcationHistory.PaymentStatus = "Success";
                                    transcationHistory.InfoMessage = "Success";
                                }
                                transcationHistory.Transactionid = entPaymentreceipt.Entities[0].Contains("hil_transactionid") ? entPaymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid") : "";
                                transcationHistory.TransactionDate = entPaymentreceipt.Entities[0].Contains("createdon") ? entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() : null;
                            }
                            string serviceName = "";
                            string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='salesorderdetail'>
                                            <attribute name='productid'/>
                                            <attribute name='salesorderdetailid'/>
                                            <order attribute='productid' descending='false'/>
                                            <filter type='and'>
                                                <condition attribute='salesorderid' operator='eq' value='{ent.Id}'/>
                                            </filter>
                                            </entity>
                                            </fetch>";
                            EntityCollection servicelineEntCol = _crmService.RetrieveMultiple(new FetchExpression(query));
                            if (servicelineEntCol.Entities.Count > 0)
                            {
                                foreach (Entity item in servicelineEntCol.Entities)
                                {
                                    Entity entProductCatalog = new Entity();
                                    if (item.Contains("productid"))
                                    {
                                        entProductCatalog = GetServiceDetails(_crmService, item.GetAttributeValue<EntityReference>("productid").Id);
                                        string tempServiceName = entProductCatalog != null ? entProductCatalog.GetAttributeValue<string>("hil_name") : "";
                                        if (tempServiceName != "")
                                        {
                                            serviceName = serviceName + tempServiceName + ", ";
                                        }
                                    }
                                }
                                if (serviceName != "")
                                    serviceName = serviceName.Substring(0, serviceName.Length - 2);
                            }
                            transcationHistory.PlanDuration = servicelineEntCol.Entities.Count.ToString() + " Services";
                            transcationHistory.ProductName = serviceName;
                            lstTranscationHistory.Add(transcationHistory);
                        }
                    }
                }
                else
                {
                    return (lstTranscationHistory, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = "D365 service unavailable."
                    });
                }
                return (lstTranscationHistory, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (lstTranscationHistory, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error : " + ex.Message.ToUpper()
                });
            }
        }
        public static Entity GetServiceDetails(IOrganizationService _service, Guid ProductId)
        {
            Entity entProductcatalog = new Entity();
            string xmlqery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                  <entity name='hil_productcatalog'>
                    <attribute name='hil_productcatalogid' />
                    <attribute name='hil_name' />
                    <attribute name='hil_plantclink' />
                    <attribute name='hil_productcode' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_productcode' operator='eq' value='{ProductId.ToString()}' />
                    </filter>
                  </entity>
                </fetch>";
            EntityCollection entitycol = _service.RetrieveMultiple(new FetchExpression(xmlqery));
            if (entitycol.Entities.Count > 0)
            {
                entProductcatalog = entitycol.Entities[0];
            }
            return entProductcatalog;
        }

        #region Models
        public class TranscationHistory
        {
            public string InvoiceId { get; set; }
            public string Transactionid { get; set; }
            public string PlanName { get; set; }
            public string ProductName { get; set; }
            public string Amount { get; set; }
            public string PlanDuration { get; set; }
            public string PaymentStatus { get; set; }
            public string TransactionDate { get; set; }
            public string InfoMessage { get; set; }
        }
        public class AMCOrdersParam
        {
            public string CustomerGuId { get; set; }
            public string SourceType { get; set; }
        }
        #endregion
    }
}
