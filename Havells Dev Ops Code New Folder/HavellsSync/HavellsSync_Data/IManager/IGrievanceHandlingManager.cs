using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.GrievanceHandling;

namespace HavellsSync_Data.IManager
{
    public interface IGrievanceHandlingManager
    {
        Task<(Response, RequestStatus)> GetComplaintCategory();
        Task<(OpenJobsResponse, RequestStatus)> GetOpenJobs(OpenJobsRequest obj);
        Task<(OpenComplaints_Response, RequestStatus)> GetOpenComplaints(OpenComplaints_Request obj);
        Task<(Complaints_Response, RequestStatus)> GetComplaints(Complaints_Request obj);
        Task<(RaiseReminder_Response, RequestStatus)> RaiseReminder(RaiseReminder_Request obj);
        Task<(CreateCase_Response, RequestStatus)> CreateCase(CreateCase_Request obj);
    }
}
