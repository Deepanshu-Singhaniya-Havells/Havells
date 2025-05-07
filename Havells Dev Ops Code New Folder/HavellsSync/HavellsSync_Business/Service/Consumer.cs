using HavellsSync_Data.IManager;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Consumer;
using HavellsSync_ModelData.ServiceAlaCarte;

namespace HavellsSync_Business.IService
{
    public class Consumer : IConsumer
    {
        private IConsumerManager _manager;
        public Consumer(IConsumerManager consumer)
        {
            Check.Argument.IsNotNull(nameof(consumer), consumer);
            _manager = consumer;
        }

        public async Task<ConsumerResponse> ConsumersAppRating(string MobileNumber, string SourceType, string Rating, string review)
        {
            return await _manager.ConsumersAppRating(MobileNumber, SourceType, Rating, review);
        }

        public async Task<InvoiceResponse> InvoiceDetails(string FromDate, string ToDate, string OrderNumber)
        {
            return await _manager.InvoiceDetails(FromDate, ToDate, OrderNumber);
        }

        public Task<RequestStatus> PriceList(List<PriceListParam> objPricelist)
        {
            return _manager.PriceList(objPricelist);
        }

        public Task<(WorkOrderResponse, RequestStatus)> GetWorkOrdersStatus(WorkOrderRequest objreq)
        {
            return _manager.GetWorkOrdersStatus(objreq);
        }

        public Task<JobStatusDTO> GetJobstatus(JobStatusDTO objreq)
        {
            return _manager.GetJobstatus(objreq);
        }

        public Task<JobRequestDTO> CreateServiceCallRequest(JobRequestDTO objreq)
        {
            return _manager.CreateServiceCallRequest(objreq);
        }

        public Task<List<JobOutput>> GetJobs(Job job)
        {
            return _manager.GetJobs(job);
        }
        public Task<List<IoTServiceCallResult>> IoTGetServiceCalls(IotServiceCall job)
        {
            return _manager.IoTGetServiceCalls(job);
        }

        public Task<List<IoTRegisteredProducts>> IoTRegisteredProducts(IoTRegisteredProducts registeredProduct)
        {
            return _manager.IoTRegisteredProducts(registeredProduct);
        }

        public Task<ReturnResult> IoTRegisterConsumer(IoT_RegisterConsumer consumer)
        {
            return _manager.IoTRegisterConsumer(consumer);
        }

        public Task<List<IoTNatureofComplaint>> IoTNatureOfComplaintByProdSubcategory(IoTNatureofComplaint natureOfComplaint)
        {
            return _manager.IoTNatureOfComplaintByProdSubcategory(natureOfComplaint);
        }

        public Task<List<NatureOfComplaint>> NatureOfComplaint(NatureOfComplaint natureOfComplaint)
        {
            return _manager.NatureOfComplaint(natureOfComplaint);
        }

        public Task<List<NatureOfComplaint>> AllNatureOfComplaints()
        {
            return _manager.AllNatureOfComplaints();
        }

        public Task<(OCLDetailsResponse, RequestStatus)> GetOCLDetails(OCLDetailsParam obj_OCL)
        {
            return _manager.GetOCLDetails(obj_OCL);
        }

    }
}
