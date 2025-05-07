using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Consumer;
using HavellsSync_ModelData.ServiceAlaCarte;

namespace HavellsSync_Business.IService
{
    public interface IConsumer
    {
       Task<ConsumerResponse> ConsumersAppRating(string MobileNumber, string SourceType, string Rating, string review);

        Task<InvoiceResponse> InvoiceDetails(string Fromdate, string Todate, string OrderNumbaer);

        Task<RequestStatus> PriceList(List<PriceListParam> objPricelist);

        Task<(WorkOrderResponse, RequestStatus)> GetWorkOrdersStatus(WorkOrderRequest objreq);

        Task<JobStatusDTO> GetJobstatus(JobStatusDTO objreq);

        Task<JobRequestDTO> CreateServiceCallRequest(JobRequestDTO objreq);

        Task<List<JobOutput>> GetJobs(Job job);
        Task<List<IoTServiceCallResult>> IoTGetServiceCalls(IotServiceCall job);

        Task<List<IoTRegisteredProducts>> IoTRegisteredProducts(IoTRegisteredProducts registeredProduct);

        Task<ReturnResult> IoTRegisterConsumer(IoT_RegisterConsumer consumer);

        Task<List<IoTNatureofComplaint>> IoTNatureOfComplaintByProdSubcategory(IoTNatureofComplaint natureOfComplaint);

        Task<List<NatureOfComplaint>> NatureOfComplaint(NatureOfComplaint natureOfComplaint);

        Task<List<NatureOfComplaint>> AllNatureOfComplaints();

        Task<(OCLDetailsResponse, RequestStatus)> GetOCLDetails(OCLDetailsParam obj_OCL);
    }
}
