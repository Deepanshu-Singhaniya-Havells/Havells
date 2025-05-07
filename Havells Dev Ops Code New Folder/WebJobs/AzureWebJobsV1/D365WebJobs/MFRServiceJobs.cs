using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Linq;

namespace D365WebJobs
{
    public class MFRServiceJobs
    {
        private IOrganizationService _service { get; set; }
        public MFRServiceJobs(IOrganizationService _service)
        {
            this._service = _service;
        }
        public JobRequestDTO CreateServiceCallRequest(JobRequestDTO _jobRequest)
        {
            JobRequestDTO _retObj = new JobRequestDTO() { StatusCode = "200", StatusDescription = "OK" };

            try
            {
                IOrganizationService service = _service;
                if (service != null)
                {
                    if (string.IsNullOrEmpty(_jobRequest.customer_mobileno) || string.IsNullOrWhiteSpace(_jobRequest.customer_mobileno))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Customer Mobile Number is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.customer_firstname) || string.IsNullOrWhiteSpace(_jobRequest.customer_firstname))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Customer First Name is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.address_line1) || string.IsNullOrWhiteSpace(_jobRequest.address_line1))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Address Line 1 is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.pincode) || string.IsNullOrWhiteSpace(_jobRequest.pincode))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Address Pincode is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.callsubtype) || string.IsNullOrWhiteSpace(_jobRequest.callsubtype))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Call Subtype is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.productsubcategory) || string.IsNullOrWhiteSpace(_jobRequest.productsubcategory))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Product Subcategory is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.natureofcomplaint) || string.IsNullOrWhiteSpace(_jobRequest.natureofcomplaint))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.callertype) || string.IsNullOrWhiteSpace(_jobRequest.callertype))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Caller Type is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.dealercode) || string.IsNullOrWhiteSpace(_jobRequest.dealercode))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Dealer Code is required." };
                    }
                    if (string.IsNullOrEmpty(_jobRequest.expecteddeliverydate) || string.IsNullOrWhiteSpace(_jobRequest.expecteddeliverydate))
                    {
                        return new JobRequestDTO { StatusCode = "204", StatusDescription = "Expected Delivery Date is required." };
                    }

                    Entity jobEntity = new Entity("hil_bulkjobsuploader");

                    jobEntity["hil_customermobileno"] = _jobRequest.customer_mobileno;
                    jobEntity["hil_customerfirstname"] = _jobRequest.customer_firstname;
                    if (!string.IsNullOrEmpty(_jobRequest.customer_lastname) && !string.IsNullOrWhiteSpace(_jobRequest.customer_lastname))
                        jobEntity["hil_customerlastname"] = _jobRequest.customer_lastname;

                    jobEntity["hil_addressline1"] = _jobRequest.address_line1;
                    if (!string.IsNullOrEmpty(_jobRequest.address_line2) && !string.IsNullOrWhiteSpace(_jobRequest.address_line2))
                        jobEntity["hil_addressline2"] = _jobRequest.address_line2;
                    if (!string.IsNullOrEmpty(_jobRequest.alternate_number) && !string.IsNullOrWhiteSpace(_jobRequest.alternate_number))
                        jobEntity["hil_alternatenumber"] = _jobRequest.alternate_number;
                    if (!string.IsNullOrEmpty(_jobRequest.area) && !string.IsNullOrWhiteSpace(_jobRequest.area))
                        jobEntity["hil_area"] = _jobRequest.area;

                    List<OptionSetDTO> callerTypeOptionSet = GetOptionSetData("hil_bulkjobsuploader", "hil_callertype", _jobRequest.callertype, service);
                    int callerTypeValue = Convert.ToInt32(callerTypeOptionSet[0].Value);
                    jobEntity["hil_callertype"] = new OptionSetValue(callerTypeValue);

                    jobEntity["hil_callsubtype"] = _jobRequest.callsubtype;
                    jobEntity["hil_dealercode"] = _jobRequest.dealercode;
                    jobEntity["hil_expecteddeliverydate"] = Convert.ToDateTime(_jobRequest.expecteddeliverydate);
                    if (!string.IsNullOrEmpty(_jobRequest.landmark) && !string.IsNullOrWhiteSpace(_jobRequest.landmark))
                        jobEntity["hil_landmark"] = _jobRequest.landmark;
                    jobEntity["hil_natureofcomplaint"] = _jobRequest.natureofcomplaint;
                    jobEntity["hil_pincode"] = _jobRequest.pincode;
                    jobEntity["hil_productsubcategory"] = _jobRequest.productsubcategory;
                    service.Create(jobEntity);
                }
                else
                {
                    _retObj = new JobRequestDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                _retObj = new JobRequestDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return _retObj;
        }

        private List<OptionSetDTO> GetOptionSetData(string entityName, string propertyName, string optionSetlabel, IOrganizationService service)
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
    public class JobRequestDTO
    {
        public string customer_firstname { get; set; }
        public string customer_lastname { get; set; }
        public string customer_mobileno { get; set; }
        public string alternate_number { get; set; }
        public string address_line1 { get; set; }
        public string address_line2 { get; set; }
        public string landmark { get; set; }
        public string pincode { get; set; }
        public string area { get; set; }
        public string callsubtype { get; set; }
        public string productsubcategory { get; set; }
        public string natureofcomplaint { get; set; }
        public string callertype { get; set; }
        public string dealercode { get; set; }
        public string expecteddeliverydate { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }

    public class JobStatusDTO
    {
        public string mobile_number { get; set; }
        public string serial_number { get; set; }
        public string product_subcategory { get; set; }
        public string call_type { get; set; }
        public string customer_complaint { get; set; }
        public string assigned_resource { get; set; }
        public string job_substatus { get; set; }
        public string technician_remarks { get; set; }
        public string closed_on { get; set; }
        public string cancel_reason { get; set; }
        public string closure_remarks { get; set; }
        public string webclosure_remarks { get; set; }
        public string status_code { get; set; }
        public string status_description { get; set; }
    }
}
