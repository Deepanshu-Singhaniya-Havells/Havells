using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ServiceAlaCarte;

namespace HavellsSync_Business.IService
{
    public interface IServiceAlaCarte
    {
        Task<(List<ProductCatagories>, RequestStatus)> GetServiceProductCategory();
        Task<(ServiceAlaCartePlanInfo, RequestStatus)> GetServiceAlaCarteList(string ProductSubCategoryId);
        Task<(ServiceOrder, RequestStatus)> PaymentRetry(PaymentRetryParam paymentRetryReq, string MobileNumber);
        Task<(List<BookedService>, RequestStatus)> GetMostBookedServices();
        Task<(ServiceOrder, RequestStatus)> CreateServiceOrder(CreateOrder objOrder);
        Task<(List<OrdersList>, RequestStatus)> GetServiceRequestList(ServiceRequest objserviceReq);
        Task<RequestStatus> RescheduleService(ReschuduleService serviceReqSech);
        Task<(ServiceRequestData, RequestStatus)> GetServiceRequestDetails(OrderDetailRequest objservice);
    }
}
