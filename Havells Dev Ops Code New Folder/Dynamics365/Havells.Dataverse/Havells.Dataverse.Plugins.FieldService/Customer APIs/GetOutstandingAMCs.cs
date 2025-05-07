using Havells.Dataverse.Plugins.FieldService.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Customer_APIs
{
    public class GetOutstandingAMCs : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            ServiceResponseData responseData = new ServiceResponseData();
            List<InvoiceInfo> lstinvoiceInfo = new List<InvoiceInfo>();
            string JsonResponse = "";
            try
            {
                if (context.InputParameters.Contains("CustomerId") && context.InputParameters["CustomerId"] is string
                    && context.InputParameters.Contains("DOP") && context.InputParameters["DOP"] is string
                    && context.InputParameters.Contains("ModelNumber") && context.InputParameters["ModelNumber"] is string
                    && context.InputParameters.Contains("SerialNumber") && context.InputParameters["SerialNumber"] is string)
                {
                    #region Validate Params
                    StringBuilder errorMessage = new StringBuilder();
                    bool IsValidRequest = true;
                    Guid CustomerId = Guid.Empty;
                    bool isValidCustomerId = Guid.TryParse(context.InputParameters["CustomerId"].ToString(), out CustomerId);
                    if (!isValidCustomerId)
                    {
                        errorMessage.AppendLine("Invalid Customer GuId.");
                        IsValidRequest = false;
                    }
                    DateTime DateOfPurchase;
                    string[] formats = { "d/MM/yyyy", "dd/MM/yyyy", "d-MM-yyyy", "dd-MM-yyyy" };
                    if (!DateTime.TryParseExact(context.InputParameters["DOP"].ToString(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOfPurchase))
                    {
                        errorMessage.AppendLine("No Content : Invalid DOP. Please Provide Date in <dd/MM/yyyy>.");
                        IsValidRequest = false;
                    }
                    if (string.IsNullOrWhiteSpace(context.InputParameters["ModelNumber"].ToString()))
                    {
                        errorMessage.AppendLine("Model Number is required.");
                        IsValidRequest = false;
                    }
                    if (string.IsNullOrWhiteSpace(context.InputParameters["SerialNumber"].ToString()))
                    {
                        errorMessage.AppendLine("Serial Number is required.");
                        IsValidRequest = false;
                    }
                    if (!IsValidRequest)
                    {
                        JsonResponse = JsonConvert.SerializeObject(new ServiceResponseData
                        {
                            result = new CustomerResult
                            {
                                ResultStatus = false,
                                ResultMessage = errorMessage.ToString(),
                            }
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    string ModelNumber = context.InputParameters["ModelNumber"].ToString();
                    string SerialNumber = context.InputParameters["SerialNumber"].ToString();
                    #endregion
                    QueryExpression query = new QueryExpression("product");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition("name", ConditionOperator.Equal, ModelNumber);
                    EntityCollection ProductColl = service.RetrieveMultiple(query);
                    if (ProductColl.Entities.Count > 0)
                    {
                        DateTime _DOPTo = DateOfPurchase.AddDays(7);
                        DateTime _DOPFrom = DateOfPurchase.AddDays(-7);
                        Guid ModelGuid = ProductColl.Entities[0].Id;
                        query = new QueryExpression("invoice");
                        query.ColumnSet = new ColumnSet("invoiceid", "createdon", "hil_productcode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);

                        FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                        filter1.AddCondition("customerid", ConditionOperator.Equal, CustomerId);
                        filter1.AddCondition("msdyn_invoicedate", ConditionOperator.OnOrAfter, _DOPFrom);
                        filter1.AddCondition("msdyn_invoicedate", ConditionOperator.OnOrBefore, _DOPTo);
                        filter1.AddCondition("hil_customerasset", ConditionOperator.Null);

                        FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                        filter2.AddCondition("hil_modelcode", ConditionOperator.Equal, ModelGuid);
                        filter2.AddCondition("hil_newserialnumber", ConditionOperator.Equal, SerialNumber);

                        query.Criteria.AddFilter(filter1);
                        query.Criteria.AddFilter(filter2);
                        EntityCollection InvoiceColl = service.RetrieveMultiple(query);

                        if (InvoiceColl.Entities.Count > 0)
                        {
                            foreach (Entity ent in InvoiceColl.Entities)
                            {
                                InvoiceInfo invoiceInfo = new InvoiceInfo();
                                invoiceInfo.InvoiceId = ent.Id;
                                invoiceInfo.CreatedOn = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                                invoiceInfo.PlanName = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;
                                lstinvoiceInfo.Add(invoiceInfo);
                            }
                            responseData.InvoiceInfo = lstinvoiceInfo;
                            responseData.result = new CustomerResult { ResultStatus = true, ResultMessage = "Success" };
                        }
                        else
                        {
                            responseData.InvoiceInfo = lstinvoiceInfo;
                            responseData.result = new CustomerResult { ResultStatus = false, ResultMessage = "No data found." };
                        }
                    }
                    else
                    {
                        responseData.result = new CustomerResult { ResultStatus = false, ResultMessage = "No Content : Please Provide Valid Model Number of Product." };
                    }
                }
                JsonResponse = JsonConvert.SerializeObject(responseData);
                context.OutputParameters["Response"] = JsonResponse;
                return;
            }
            catch (Exception ex)
            {
                JsonResponse = JsonConvert.SerializeObject(new ServiceResponseData
                {
                    result = new CustomerResult
                    {
                        ResultStatus = false,
                        ResultMessage = ex.Message
                    }
                });
                context.OutputParameters["Response"] = JsonResponse;
                return;
            }
        }
    }
}
