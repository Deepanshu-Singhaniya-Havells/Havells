using System.Collections.Generic;

namespace HavellsSync_ModelData.GrievanceHandling
{
    public class CreateCase_Request
    {
        public string CustomerGuid { get; set; }
        public string AddressGuid { get; set; }
        public string ServiceRequestGuid { get; set; }
        public string ComplaintCategoryId { get; set; }
        public string ComplaintTitle { get; set; }
        public string ComplaintDescription { get; set; }
        public string Attachment { get; set; }
        public string FileName { get; set; }

    }

    public class CreateCase_ServiceRequest
    {
        public string ComplaintId { get; set; }

        public string ComplaintGuid { get; set; }

        public string Attachment_URL { get; set; }

    }
    public class CreateCase_Response
    {
       public List<CreateCase_ServiceRequest> ServiceRequests { get; set; }
    }

    public class RaiseReminder_Request
    {
        public string ComplaintGuid { get; set; }
        public string ReminderRemarks { get; set; }
    }

    public class RaiseReminder_Response
    {
        public bool Status { get; set; }

        public string Message { get; set; }
    }


    public class Complaints_Request
    {
        public string CustomerGuid { get; set; }
    }
    public class Complaints_Complaint
    {
        public string ComplaintId { get; set; }

        public string CompliantGuid { get; set; }

        public string Status { get; set; }

        public string ComplaintCategory { get; set; }

        public string Created_On { get; set; }

        public string Service_RequestID { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public string AddressLine1 { get; set; }

        public string AddressLine2 { get; set; }

        public string AddressCity { get; set; }

        public string AddressState { get; set; }

        public string AddressPinCode { get; set; }

        public string Attachment_URL { get; set; }

        public string ReminderRemarks { get; set; }

        public bool IsReminderSet { get; set; }
        public string Reminder_RaisedOn { get; set; }
        public string EscalatedOn { get; set; }
        public string Escalation_Level { get; set; }

        public string AssignedOn { get; set; }

        public string FirstResponseSentOn { get; set; }

        public string ResolvedOn { get; set; }


    }
    public class Complaints_Response
    {
        public string Error { get; set; }
        public List<Complaints_Complaint> Complaints { get; set; }
    }



    public class OpenComplaints_Request
    {
        public string CustomerGuid { get; set; }
    }

    public class OpenComplaints_Complaint
    {
        public string ComplaintId { get; set; }

        public string CompliantGuid { get; set; }

        public string Status { get; set; }

        public string ComplaintCategory { get; set; }

        public string Created_On { get; set; }

        public string Service_RequestID { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public string AddressLine1 { get; set; }

        public string AddressLine2 { get; set; }

        public string AddressCity { get; set; }

        public string AddressState { get; set; }

        public string AddressPinCode { get; set; }

        public string Attachment_URL { get; set; }

        public string ReminderRemarks { get; set; }

        public bool IsReminderSet { get; set; }
        public string Reminder_RaisedOn { get; set; }
        public string EscalatedOn { get; set; }
        public string Escalation_Level { get; set; }
        public string AssignedOn { get; set; }

        public string FirstResponseSentOn { get; set; }
        public string ResolvedOn { get; set; }
       

    }
    public class OpenComplaints_Response
    {
        public string Error { get; set; }
        public List<OpenComplaints_Complaint> Complaints { get; set; }
    }
    public class OpenJobsRequest
    {
        public string CustomerGuid { get; set; }
    }


    public class OpenJobs_ServiceRequest
    {
        public string ServiceRequestId { get; set; }

        public string ServiceRequestGUID { get; set; }

        public string Status { get; set; }

        public string Created_On { get; set; }

        public string Product_SubCategory { get; set; }

    }
    public class OpenJobsResponse
    {
        public string Error { get; set; }
        public List<OpenJobs_ServiceRequest> ServiceRequests { get; set; }
    }
    public class ComplaintCategory
    {
        public string CategoryGUID { get; set; }

        public string CategoryName { get; set; }

        public string Department { get; set; }

    }
    public class Response
    {
        public List<ComplaintCategory> ComplaintCategories { get; set; }
    }

    public class param<T>
    {
        public T data { get; set; }
    }
    public class SourceTypes
    {
        public string SourceType { get; set; }

    }

   

}