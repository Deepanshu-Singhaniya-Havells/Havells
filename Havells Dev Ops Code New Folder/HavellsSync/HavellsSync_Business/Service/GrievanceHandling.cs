using HavellsSync_Business.IService;
using HavellsSync_Data.IManager;
using HavellsSync_Data.Manager;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.GrievanceHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Business.IService
{
    public class GrievanceHandling : IGrievanceHandling
    {
        private IGrievanceHandlingManager _manager;
        public GrievanceHandling(IGrievanceHandlingManager Grievance)
        {
            Check.Argument.IsNotNull(nameof(Grievance), Grievance);
            _manager = Grievance;
        }
        public async Task<(Response, RequestStatus)> GetComplaintCategory()
        {
            return await _manager.GetComplaintCategory();
        }
        public async Task<(OpenJobsResponse, RequestStatus)> GetOpenJobs(OpenJobsRequest obj)
        {
            return await _manager.GetOpenJobs(obj);
        }
        public async Task<(OpenComplaints_Response, RequestStatus)> GetOpenComplaints(OpenComplaints_Request obj)
        {
            return await _manager.GetOpenComplaints(obj);
        }
        public async Task<(Complaints_Response, RequestStatus)> GetComplaints(Complaints_Request obj)
        {
            return await _manager.GetComplaints(obj); // all complaints
        }
        public async Task<(RaiseReminder_Response, RequestStatus)> RaiseReminder(RaiseReminder_Request obj)
        {
            return await _manager.RaiseReminder(obj);
        }
        public async Task<(CreateCase_Response, RequestStatus)> CreateCase(CreateCase_Request obj)
        {
            return await _manager.CreateCase(obj);
        }
    }
}
