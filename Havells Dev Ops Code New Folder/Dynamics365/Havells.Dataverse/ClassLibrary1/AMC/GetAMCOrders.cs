using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class GetAMCOrders : IPlugin
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
                string msg = string.IsNullOrWhiteSpace(CustomerGuid) ? "Customer guid is required." : "Invalid Customer guid.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (string.IsNullOrWhiteSpace(SourceType))
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Source type is required." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (SourceType != "6")
            {
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Please enter valid Source Type." });
                context.OutputParameters["data"] = responnse;
                return;
            }
            AMCOrdersParam objparam = new AMCOrdersParam();
            objparam.CustomerGuId = CustomerGuid;
            objparam.SourceType = SourceType;
            var response = GetAMCOrder(service, objparam, LoginUserId);
            dynamic result;

            if (response.Item2.StatusCode == (int)HttpStatusCode.OK)
                result = System.Text.Json.JsonSerializer.Serialize(response.Item1);
            else
                result = System.Text.Json.JsonSerializer.Serialize(response.Item2);
            context.OutputParameters["data"] = result;
            return;
        }
        public (AMCOrdersListRes, RequestStatus) GetAMCOrder(IOrganizationService _crmService, AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            AMCOrdersListRes lstAMCOrdersRes = new AMCOrdersListRes();
            lstAMCOrdersRes.AMCOrders = new List<AMCOrders>();
            StatusRequest reqParm = new StatusRequest();
            string TxnId = string.Empty;
            Guid CustomerGuId = Guid.Empty;
            try
            {
                if (_crmService != null)
                {
                    try
                    {
                        CustomerGuId = new Guid(AMCOrdersParam.CustomerGuId);
                    }
                    catch (Exception)
                    {
                        return (lstAMCOrdersRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid customer Guid."
                        });
                    }
                    Entity entity = _crmService.Retrieve("contact", CustomerGuId, new ColumnSet(false));

                    if (entity == null)
                    {
                        return (lstAMCOrdersRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.NotFound,
                            Message = "Requested resource does not exist"
                        });
                    }
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='salesorder'>
                                        <attribute name='name' />
                                        <attribute name='customerid' />
                                        <attribute name='statuscode' />
                                        <attribute name='totalamount' />
                                        <attribute name='salesorderid' />
                                        <attribute name='hil_paymentstatus' />
                                        <attribute name='createdon' />
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                          <condition attribute='customerid' operator='eq' value='{AMCOrdersParam.CustomerGuId}' />
                                          <condition attribute='hil_ordertype' operator='eq' value='{{1F9E3353-0769-EF11-A670-0022486E4ABB}}' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";

                    EntityCollection entityColl = _crmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entityColl.Entities.Count > 0)
                    {
                        foreach (Entity ent in entityColl.Entities)
                        {
                            AMCOrders amcOrdersRes = new AMCOrders();
                            amcOrdersRes.InvoiceId = ent.Id.ToString();
                            amcOrdersRes.InvoiceDate = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString("dd/MM/yyyy");
                            amcOrdersRes.InvoiceValue = ent.Contains("totalamount") ? decimal.Round(ent.GetAttributeValue<Money>("totalamount").Value, 2).ToString() : null;
                            //amcOrdersRes.InvoiceDescription =ent.Contains("description") ? ent.GetAttributeValue<string>("description").ToString() : null;
                            int PaymentStatus = ent.Contains("hil_paymentstatus") ? ent.GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 0;
                            if (PaymentStatus == 1 || PaymentStatus == 3)
                            {
                                amcOrdersRes.PaymentStatus = "Pending";
                            }
                            else if (PaymentStatus == 2)
                            {
                                amcOrdersRes.PaymentStatus = "Success";
                            }
                            else
                            {
                                amcOrdersRes.PaymentStatus = "Failed";
                            }
                            //fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                            //            <entity name='hil_paymentreceipt'>
                            //            <attribute name='hil_paymentreceiptid' />
                            //            <attribute name='hil_transactionid' />
                            //            <attribute name='hil_paymentstatus' />
                            //            <attribute name='hil_tokenexpireson' />
                            //            <order attribute='createdon' descending='true' />
                            //            <filter type='and'>
                            //                <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                            //            </filter>
                            //            </entity>
                            //            </fetch>";
                            //EntityCollection _entColPayment = _crmService.RetrieveMultiple(new FetchExpression(fetchXml));
                            //if (_entColPayment.Entities.Count > 0)
                            //{
                            //    int paymentstatus = _entColPayment.Entities[0].Contains("hil_paymentstatus") ? _entColPayment.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 1;

                            //    if (paymentstatus == 1 || paymentstatus == 3)//Payment Initiated || In Progress
                            //    {
                            //        TxnId = _entColPayment.Entities[0].GetAttributeValue<string>("hil_transactionid");
                            //        string Status = CommonMethods.getTransactionStatus(_crmService, _entColPayment.Entities[0].Id, TxnId, ent.Id, _Paymenturlkey);

                            //        if (Status == "Success" || Status == "Failed" || Status == "Pending")
                            //        {
                            //            amcOrdersRes.PaymentStatus = Status;
                            //        }
                            //        else
                            //        {
                            //            return (lstAMCOrdersRes, new RequestStatus()
                            //            {
                            //                StatusCode = (int)HttpStatusCode.BadRequest,
                            //                Message = CommonMessage.InternalServerErrorMsg + Status
                            //            });
                            //        }
                            //    }
                            //    else if (paymentstatus == 4)//Paid
                            //    {
                            //        amcOrdersRes.PaymentStatus = "Success";
                            //    }
                            //    else
                            //    {
                            //        amcOrdersRes.PaymentStatus = "Failed";
                            //    }
                            //}
                            lstAMCOrdersRes.AMCOrders.Add(amcOrdersRes);
                        }
                        lstAMCOrdersRes.StatusCode = (int)HttpStatusCode.OK;
                    }
                }
                else
                {
                    return (lstAMCOrdersRes, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = "D365 service unavailable."
                    });
                }
                return (lstAMCOrdersRes, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (lstAMCOrdersRes, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error : " + ex.Message.ToUpper()
                });
            }
        }

        #region Get AMC Orders
        public class AMCOrdersParam
        {
            public string CustomerGuId { get; set; }
            public string SourceType { get; set; }
        }
        public class StatusRequest
        {
            public string PROJECT { get; set; }
            public string command { get; set; }
            public string var1 { get; set; }
        }
        public class AMCOrdersListRes
        {
            public int StatusCode { get; set; }
            public List<AMCOrders> AMCOrders { get; set; }
        }
        public class AMCOrders
        {
            public string InvoiceId { get; set; }
            public string InvoiceDate { get; set; }
            public string InvoiceValue { get; set; }
            public string PaymentStatus { get; set; }
            public string InvoiceDescription { get; set; }
        }
        public class RequestStatus
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
        }

        #endregion
    }
}
