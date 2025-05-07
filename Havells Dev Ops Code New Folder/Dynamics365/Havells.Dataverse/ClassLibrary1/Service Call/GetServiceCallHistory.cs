using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Service_Call
{
    public class GetServiceCallHistory : IPlugin
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
            Guid CustomerGuid = Guid.Empty;
            List<IoTServiceCallResult> lstServiceCallResult = new List<IoTServiceCallResult>();

            if (context.InputParameters.Contains("CustomerGuid"))
            {
                bool isValidCustomerId = Guid.TryParse(Convert.ToString(context.InputParameters["CustomerGuid"]), out CustomerGuid);
                if (!isValidCustomerId)
                {
                    lstServiceCallResult.Add(new IoTServiceCallResult()
                    {
                        StatusCode = "204",
                        StatusDescription = "Invalid Customer GUID."
                    });
                    JsonResponse = JsonSerializer.Serialize(lstServiceCallResult);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                JsonResponse = JsonSerializer.Serialize(Process(service, CustomerGuid));
            }
            else
            {
                lstServiceCallResult.Add(new IoTServiceCallResult()
                {
                    StatusCode = "204",
                    StatusDescription = "Customer GUID is required."
                });
                JsonResponse = JsonSerializer.Serialize(lstServiceCallResult);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            context.OutputParameters["data"] = JsonResponse;
        }
        public List<IoTServiceCallResult> Process(IOrganizationService service, Guid CustomerGuid)
        {
            List<IoTServiceCallResult> jobList = new List<IoTServiceCallResult>();
            IoTServiceCallResult objJobOutput;
            try
            {
                if (service != null)
                {
                    QueryExpression query = new QueryExpression()
                    {
                        EntityName = "msdyn_workorder",
                        ColumnSet = new ColumnSet(true)
                    };
                    FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                    filterExpression.Conditions.Add(new ConditionExpression("hil_customerref", ConditionOperator.Equal, CustomerGuid));
                    query.Criteria.AddFilter(filterExpression);
                    query.AddOrder("createdon", OrderType.Descending);

                    EntityCollection collection = service.RetrieveMultiple(query);

                    if (collection.Entities != null && collection.Entities.Count > 0)
                    {
                        foreach (Entity item in collection.Entities)
                        {
                            IoTServiceCallResult jobObj = new IoTServiceCallResult();
                            if (item.Attributes.Contains("msdyn_name"))
                            {
                                jobObj.JobId = item.GetAttributeValue<string>("msdyn_name");
                            }
                            if (item.Attributes.Contains("msdyn_name"))
                            {
                                jobObj.JobGuid = item.Id;
                            }
                            if (item.Attributes.Contains("hil_callsubtype"))
                            {
                                jobObj.CallSubType = item.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                            }
                            if (item.Attributes.Contains("createdon"))
                            {
                                jobObj.JobLoggedon = item.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                            }
                            if (item.Attributes.Contains("msdyn_substatus"))
                            {
                                jobObj.JobStatus = item.GetAttributeValue<EntityReference>("msdyn_substatus").Name;
                            }
                            if (item.Attributes.Contains("hil_owneraccount"))
                            {
                                jobObj.JobAssignedTo = item.GetAttributeValue<EntityReference>("hil_owneraccount").Name;
                            }
                            if (item.Attributes.Contains("msdyn_customerasset"))
                            {
                                jobObj.CustomerAsset = item.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                            }
                            if (item.Attributes.Contains("hil_productcategory"))
                            {
                                jobObj.ProductCategory = item.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                            }
                            if (item.Attributes.Contains("hil_natureofcomplaint"))
                            {
                                jobObj.NatureOfComplaint = item.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Name;
                            }
                            if (item.Attributes.Contains("hil_jobclosuredon"))
                            {
                                jobObj.JobClosedOn = item.GetAttributeValue<DateTime>("hil_jobclosuredon").AddMinutes(330).ToString();
                            }
                            if (item.Attributes.Contains("hil_customerref"))
                            {
                                jobObj.CustomerName = item.GetAttributeValue<EntityReference>("hil_customerref").Name;
                            }
                            if (item.Attributes.Contains("hil_fulladdress"))
                            {
                                jobObj.ServiceAddress = item.GetAttributeValue<string>("hil_fulladdress");
                            }
                            if (item.Attributes.Contains("msdyn_customerasset"))
                            {
                                Entity ec = service.Retrieve("msdyn_customerasset", item.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_modelname"));
                                if (ec != null)
                                {
                                    jobObj.Product = ec.GetAttributeValue<string>("hil_modelname");
                                }
                            }
                            if (item.Attributes.Contains("hil_customercomplaintdescription"))
                            {
                                jobObj.ChiefComplaint = item.GetAttributeValue<string>("hil_customercomplaintdescription");
                            }
                            if (item.Attributes.Contains("hil_preferredtime"))
                            {
                                jobObj.PreferredPartOfDay = item.GetAttributeValue<OptionSetValue>("hil_preferredtime").Value;
                                jobObj.PreferredPartOfDayName = item.FormattedValues["hil_preferredtime"].ToString();
                            }
                            if (item.Attributes.Contains("hil_preferreddate"))
                            {
                                jobObj.PreferredDate = item.GetAttributeValue<DateTime>("hil_preferreddate").AddMinutes(330).ToShortDateString();
                            }
                            jobObj.StatusCode = "200";
                            jobObj.StatusDescription = "OK";
                            jobList.Add(jobObj);
                        }
                    }
                    return jobList;
                }
                else
                {
                    objJobOutput = new IoTServiceCallResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    jobList.Add(objJobOutput);
                }
            }
            catch (Exception ex)
            {
                objJobOutput = new IoTServiceCallResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                jobList.Add(objJobOutput);
            }
            return jobList;
        }
    }

    public class IoTServiceCallResult
    {
        public string JobId { get; set; }
        public Guid JobGuid { get; set; }
        public string CallSubType { get; set; }
        public string JobLoggedon { get; set; }
        public string JobStatus { get; set; }
        public string JobAssignedTo { get; set; }
        public string CustomerAsset { get; set; }
        public string ProductCategory { get; set; }
        public string NatureOfComplaint { get; set; }
        public string JobClosedOn { get; set; }
        public string CustomerName { get; set; }
        public string ServiceAddress { get; set; }
        public string Product { get; set; }
        public string ChiefComplaint { get; set; }
        public string PreferredDate { get; set; }
        public int PreferredPartOfDay { get; set; }
        public string PreferredPartOfDayName { get; set; }
        public string StatusCode { get; set; } = "204";
        public string StatusDescription { get; set; } = "No record found";
    }
}
