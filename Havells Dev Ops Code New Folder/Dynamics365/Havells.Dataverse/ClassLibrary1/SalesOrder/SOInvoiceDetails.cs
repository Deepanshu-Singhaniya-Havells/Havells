using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class SOInvoiceDetails : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string JsonResponse = "";
            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            reqData reqdata = JsonSerializer.Deserialize<reqData>(jsonString);
          
            if (string.IsNullOrWhiteSpace(reqdata.FromDate) && string.IsNullOrWhiteSpace(reqdata.ToDate) && string.IsNullOrWhiteSpace(reqdata.OrderNumber))
            {
                JsonResponse = JsonSerializer.Serialize(new InvoiceResponse
                {
                    Status = false,
                    Message = "Invalid Request."
                });
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            if (string.IsNullOrWhiteSpace(reqdata.OrderNumber))
            {
                if (!APValidate.IsvalidDate(reqdata.FromDate))
                {
                    JsonResponse = JsonSerializer.Serialize(new InvoiceResponse
                    {
                        Status = false,
                        Message = "Require a valid FromDate."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else if (!APValidate.IsvalidDate(reqdata.ToDate))
                {
                    JsonResponse = JsonSerializer.Serialize(new InvoiceResponse
                    {
                        Status = false,
                        Message = "Require a valid ToDate."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                string DateValidationMessage = APValidate.ValidateTwoDates(Convert.ToDateTime(reqdata.ToDate), Convert.ToDateTime(reqdata.FromDate));
                if (!string.IsNullOrWhiteSpace(DateValidationMessage))
                {
                    JsonResponse = JsonSerializer.Serialize(new InvoiceResponse
                    {
                        Status = false,
                        Message = DateValidationMessage
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            JsonResponse = JsonSerializer.Serialize(GetOrderDetails(service, reqdata.FromDate, reqdata.ToDate, reqdata.OrderNumber));
            context.OutputParameters["data"] = JsonResponse;
        }
        public InvoiceResponse GetOrderDetails(IOrganizationService _CrmService, string FromDate, string ToDate, string OrderNumber)
        {
            InvoiceResponse InvResponse = new InvoiceResponse();
            try
            {
                if (!string.IsNullOrWhiteSpace(OrderNumber))
                {
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='salesorder'>
                                        <attribute name='name' />
                                        <attribute name='salesorderid' />
                                        <attribute name='hil_sourcereferencecode' />
                                        <attribute name='hil_sapsyncmessage' />
                                        <attribute name='hil_sapsonumber' />
                                        <attribute name='hil_sapinvoicenumber' />
                                        <order attribute='name' descending='false' />
                                        <filter type='and'>
                                          <filter type='or'>
                                            <condition attribute='hil_sellingsource' operator='eq' value='{{608E899B-A8A3-ED11-AAD1-6045BDAD27A7}}' />
                                            <condition attribute='hil_source' operator='eq' value='22' />
                                          </filter>
                                          <condition attribute='hil_sourcereferencecode' operator='eq' value='{OrderNumber}' />
                                        </filter>
                                      </entity>
                                    </fetch>";

                    EntityCollection entColl = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entColl.Entities.Count > 0)
                    {
                        List<InvoiceDetailsResponse> InvcResponse = new List<InvoiceDetailsResponse>();
                        foreach (Entity ent in entColl.Entities)
                        {
                            InvoiceDetailsResponse IDR = new InvoiceDetailsResponse();
                            IDR.OrderNumber = ent.Contains("hil_sourcereferencecode") ? ent.GetAttributeValue<string>("hil_sourcereferencecode") : "";
                            IDR.SAPOrderNumber = ent.Contains("hil_sapsonumber") ? ent.GetAttributeValue<string>("hil_sapsonumber") : null;
                            IDR.SAPInvoiceNumber = ent.Contains("hil_sapinvoicenumber") ? ent.GetAttributeValue<string>("hil_sapinvoicenumber") : null;
                            IDR.SAPSyncMessage = ent.Contains("hil_sapsyncmessage") ? ent.GetAttributeValue<string>("hil_sapsyncmessage") : null;
                            InvcResponse.Add(IDR);
                        }
                        InvResponse.Data = InvcResponse;
                        InvResponse.Message = "Success";
                        InvResponse.Status = true;
                    }
                }
                else if (!string.IsNullOrWhiteSpace(FromDate) && !string.IsNullOrWhiteSpace(ToDate))
                {
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='salesorder'>
                                        <attribute name='name' />
                                        <attribute name='salesorderid' />
                                        <attribute name='hil_sourcereferencecode' />
                                        <attribute name='hil_sapsyncmessage' />
                                        <attribute name='hil_sapsonumber' />
                                        <attribute name='hil_sapinvoicenumber' />
                                        <order attribute='name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='createdon' operator='on-or-after' value='{FromDate}' />
                                          <condition attribute='createdon' operator='on-or-before' value='{ToDate}' />
                                           <filter type='or'>
                                            <condition attribute='hil_sellingsource' operator='eq' value='{{608E899B-A8A3-ED11-AAD1-6045BDAD27A7}}' />
                                            <condition attribute='hil_source' operator='eq' value='22' />
                                          </filter>
                                        </filter>
                                      </entity>
                                    </fetch>";

                    EntityCollection Info1 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (Info1.Entities.Count > 0)
                    {
                        List<InvoiceDetailsResponse> InvcResponse = new List<InvoiceDetailsResponse>();
                        foreach (Entity ent in Info1.Entities)
                        {
                            InvoiceDetailsResponse IDR = new InvoiceDetailsResponse();
                            IDR.OrderNumber = ent.Contains("hil_sourcereferencecode") ? ent.GetAttributeValue<string>("hil_sourcereferencecode") : "";
                            IDR.SAPOrderNumber = ent.Contains("hil_sapsonumber") ? ent.GetAttributeValue<string>("hil_sapsonumber") : null;
                            IDR.SAPInvoiceNumber = ent.Contains("hil_sapinvoicenumber") ? ent.GetAttributeValue<string>("hil_sapinvoicenumber") : null;
                            IDR.SAPSyncMessage = ent.Contains("hil_sapsyncmessage") ? ent.GetAttributeValue<string>("hil_sapsyncmessage") : null;
                            InvcResponse.Add(IDR);
                        }
                        InvResponse.Data = InvcResponse;
                        InvResponse.Message = "Success";
                        InvResponse.Status = true;
                    }
                }
            }
            catch (Exception ex)
            {
                InvResponse.Message = ex.Message;
                InvResponse.Status = false;
            }
            return InvResponse;
        }
    }
    public class InvoiceResponse
    {
        public List<InvoiceDetailsResponse> Data { set; get; }
        public bool Status { get; set; }
        public string Message { get; set; }

    }
    public class reqData
    {
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public string OrderNumber { get; set; }
    }

    public class InvoiceDetailsResponse
    {
        public string OrderNumber { get; set; }
        public string SAPInvoiceNumber { get; set; }
        public string SAPOrderNumber { get; set; }
        public string SAPSyncMessage { get; set; }
    }
}
