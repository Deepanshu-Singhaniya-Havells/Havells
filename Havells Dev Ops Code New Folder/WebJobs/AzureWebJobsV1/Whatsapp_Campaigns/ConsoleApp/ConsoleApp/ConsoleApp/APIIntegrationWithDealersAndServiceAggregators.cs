using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace OMSMail
{
    class APIIntegrationWithDealersAndServiceAggregators
    {
        private readonly IOrganizationService service;

        public APIIntegrationWithDealersAndServiceAggregators(IOrganizationService _service)
        {
            service = _service;
        }
        public JobStatusDetails GetJobstatus(ARequestData reqestData)
        {
            JobStatusDetails responseData = new JobStatusDetails();
            try
            {
                if (string.IsNullOrWhiteSpace(reqestData.JobId))
                {
                    responseData.result = new Result { ResultStatus = false, ResultMessage = "No Content : Invalid JobId." };
                    return responseData;
                }
                string FetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='msdyn_workorder'>
                            <attribute name='msdyn_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='hil_customerref' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='msdyn_workorderid' />
                            <attribute name='hil_webclosureremarks' />
                            <attribute name='hil_typeofassignee' />
                            <attribute name='msdyn_substatus' />
                            <attribute name='hil_productcategory' />
                            <attribute name='ownerid' />
                            <attribute name='hil_mobilenumber' />
                            <attribute name='hil_jobcancelreason' />
                            <attribute name='hil_customercomplaintdescription' />
                            <attribute name='hil_closureremarks' />
                            <attribute name='msdyn_timeclosed' />
                            <attribute name='msdyn_customerasset' />
                            <order attribute='msdyn_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='msdyn_name' operator='eq' value='{reqestData.JobId}' />
                            </filter>
                          </entity>
                        </fetch>";

                EntityCollection JobDetailsColl = service.RetrieveMultiple(new FetchExpression(FetchQuery));
                if (JobDetailsColl.Entities.Count > 0)
                {
                    foreach (Entity entity in JobDetailsColl.Entities)
                    {
                        responseData.MobileNumber = entity.Contains("hil_mobilenumber") ? entity.GetAttributeValue<string>("hil_mobilenumber") : "";
                        responseData.JobID = entity.Contains("msdyn_name") ? entity.GetAttributeValue<string>("msdyn_name") : "";
                        responseData.ProductSerialNumber = entity.Contains("msdyn_customerasset") ? entity.GetAttributeValue<EntityReference>("msdyn_customerasset").Name : "";
                        responseData.ProductSubCategory = entity.Contains("hil_productsubcategory") ? entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : "";
                        responseData.CallSubType = entity.Contains("hil_callsubtype") ? entity.GetAttributeValue<EntityReference>("hil_callsubtype").Name : "";
                        responseData.CustomerComplaintDescription = entity.Contains("hil_customercomplaintdescription") ? entity.GetAttributeValue<string>("hil_customercomplaintdescription") : "";
                        responseData.AssignedTo = entity.Contains("ownerid") ? entity.GetAttributeValue<EntityReference>("ownerid").Name : "";
                        responseData.JobSubStatus = entity.Contains("msdyn_substatus") ? entity.GetAttributeValue<EntityReference>("msdyn_substatus").Name : "";
                        responseData.ClosedOn = entity.Contains("msdyn_timeclosed") ? entity.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString() : "";
                        if (entity.Attributes.Contains("hil_jobcancelreason"))
                        {
                            if (entity.FormattedValues.Contains("hil_jobcancelreason"))
                                responseData.JobCancelReason = entity.FormattedValues["hil_jobcancelreason"];
                        }
                        responseData.WebClosureRemarks = entity.Contains("hil_webclosureremarks") ? entity.GetAttributeValue<string>("hil_webclosureremarks") : "";
                        responseData.ClosureRemarks = entity.Contains("hil_closureremarks") ? entity.GetAttributeValue<string>("hil_closureremarks") : "";
                        string IncFetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='msdyn_workorderproduct'>
                                                    <attribute name='msdyn_product' />
                                                    <filter type='and'>
                                                      <condition attribute='msdyn_workorder' operator='eq' value='{entity.Id}' />
                                                      <condition attribute='statecode' operator='eq' value='0' />
                                                      <condition attribute='hil_markused' operator='eq' value='1' />
                                                    </filter>
                                                    <link-entity name='msdyn_workorderincident' from='msdyn_workorderincidentid' to='msdyn_workorderincident' visible='false' link-type='outer' alias='wi'>
                                                      <attribute name='msdyn_description' />
                                                    </link-entity>
                                                  </entity>
                                                </fetch>";
                        EntityCollection JobIncDetailsColl = service.RetrieveMultiple(new FetchExpression(IncFetchQuery));
                        StringBuilder SparePart = new StringBuilder();
                        if (JobIncDetailsColl.Entities.Count > 0)
                        {
                            responseData.TechnicianRemarks = JobIncDetailsColl.Entities[0].Contains("wi.msdyn_description") ? JobIncDetailsColl.Entities[0].GetAttributeValue<AliasedValue>("wi.msdyn_description").Value.ToString() : "";
                            foreach (Entity ent in JobIncDetailsColl.Entities)
                            {
                                SparePart.Append((ent.Contains("msdyn_product") ? ent.GetAttributeValue<EntityReference>("msdyn_product").Name : "") + ", ");
                            }
                            responseData.SparePart = SparePart.Length > 2 ? SparePart.ToString().Substring(0, SparePart.Length - 2) : "";
                        }
                        responseData.result = new Result { ResultStatus = true, ResultMessage = "Success" };
                    }
                }
            }
            catch (Exception ex)
            {
                responseData.result = new Result { ResultStatus = false, ResultMessage = ex.Message };
                return responseData;
            }
            Console.WriteLine(JsonConvert.SerializeObject(responseData));
            return responseData;
        }

        public void CreateBulkJobs(List<JobInfo> jobs)
        {
            try
            {
                if (jobs != null)
                {
                    foreach (var job in jobs)
                    {
                        Entity jobEntity = new Entity("hil_bulkjobsuploader");
                        jobEntity["hil_customermobileno"] = job.CustomerMobileNo;
                        jobEntity["hil_customerfirstname"] = job.CustomerFirstName;
                        jobEntity["hil_customerlastname"] = job.CustomerLastName;
                        jobEntity["hil_addressline1"] = job.AddressLine1;
                        jobEntity["hil_addressline2"] = job.AddressLine2;
                        jobEntity["hil_alternatenumber"] = job.AlternateNumber;
                        jobEntity["hil_area"] = job.Area;
                        Task<List<OptionSetDTO>> callerTypeOptionSet = GetOptionSetData("hil_bulkjobsuploader", "hil_callertype", job.CallerType);
                        int callerTypeValue = Convert.ToInt32(callerTypeOptionSet.Result[0].Value);
                        jobEntity["hil_callertype"] = new OptionSetValue(callerTypeValue);
                        jobEntity["hil_callsubtype"] = job.CallSubType;
                        jobEntity["hil_dealercode"] = job.DealerNameOrCode;
                        jobEntity["hil_expecteddeliverydate"] = Convert.ToDateTime(job.ExpectedDeliveryDate);
                        jobEntity["hil_jobstatus"] = job.JobStatus;
                        jobEntity["hil_landmark"] = job.Landmark;
                        jobEntity["hil_natureofcomplaint"] = job.NatureofComplaint;
                        jobEntity["hil_pincode"] = job.PinCode;
                        jobEntity["hil_productsubcategory"] = job.ProductSubCategory;
                        jobEntity["hil_salutation"] = job.Salutation;
                        service.Create(jobEntity);
                        Console.WriteLine("Created Successfully!");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private async Task<List<OptionSetDTO>> GetOptionSetData(string entityName, string propertyName, string optionSetlabel)
        {
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = propertyName,
                RetrieveAsIfPublished = false
            };

            RetrieveAttributeResponse response = service.Execute(attributeRequest) as RetrieveAttributeResponse;

            EnumAttributeMetadata attributeData = (EnumAttributeMetadata)response.AttributeMetadata;

            var optionList = (from option in attributeData.OptionSet.Options
                              where option.Label.UserLocalizedLabel.Label.Equals(optionSetlabel)
                              select new OptionSetDTO(
                                option.Value,
                                option.Label.UserLocalizedLabel.Label)).ToList();
            return optionList;
        }

    }
    public class ARequestData
    {
        public string JobId { get; set; }
    }
    public class Result
    {
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
    }
    public class JobStatusDetails
    {
        public string MobileNumber { get; set; }
        public string JobID { get; set; }
        public string ProductSerialNumber { get; set; }
        public string ProductSubCategory { get; set; }
        public string CallSubType { get; set; }
        public string CustomerComplaintDescription { get; set; }
        public string AssignedTo { get; set; }
        public string JobSubStatus { get; set; }
        public string TechnicianRemarks { get; set; }
        public string SparePart { get; set; }
        public string ClosedOn { get; set; }
        public string JobCancelReason { get; set; }
        public string WebClosureRemarks { get; set; }
        public string ClosureRemarks { get; set; }
        public Result result { get; set; }

    }
    public class OptionSetDTO
    {
        public OptionSetDTO(int? _value, string _name)
        {
            Value = _value;
            Name = _name;
        }
        public int? Value { get; set; }
        public string Name { get; set; }
    }
    public class JobInfo
    {
        public string CustomerMobileNo { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string AlternateNumber { get; set; }
        public string Area { get; set; }
        public string CallerType { get; set; }
        public string CallSubType { get; set; }
        public string CustomerFirstName { get; set; }
        public string CustomerLastName { get; set; }
        public string DealerNameOrCode { get; set; }
        public string ExpectedDeliveryDate { get; set; }
        public bool JobStatus { get; set; }
        public string Landmark { get; set; }
        public string NatureofComplaint { get; set; }
        public string PinCode { get; set; }
        public string ProductSubCategory { get; set; }
        public string Salutation { get; set; }
    }

}
