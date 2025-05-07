using HavellsSync_Business.IService;
using HavellsSync_Data.IManager;
using HavellsSync_Data.IServiceAlaCarte;
using HavellsSync_Data.Manager;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ServiceAlaCarte;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Business.IService
{
    public class ServiceAlaCarte : IServiceAlaCarte
    {
        private IServiceAlaCarteManager _manager;
        public ServiceAlaCarte(IServiceAlaCarteManager AlaCarte)
        {
            Check.Argument.IsNotNull(nameof(AlaCarte), AlaCarte);
            _manager = AlaCarte;
        }
        public async Task<(List<ProductCatagories>, RequestStatus)> GetServiceProductCategory()
        {
            return await _manager.GetServiceProductCategory();
        }
        public async Task<(ServiceAlaCartePlanInfo, RequestStatus)> GetServiceAlaCarteList(string ProductSubCategoryId)
        {
            return await _manager.GetServiceAlaCarteList(ProductSubCategoryId);
        }
        public async Task<(ServiceOrder, RequestStatus)> PaymentRetry(PaymentRetryParam paymentRetryReq, string MobileNumber)
        {
            return await _manager.PaymentRetry(paymentRetryReq, MobileNumber);
        }
        public Task<(List<BookedService>, RequestStatus)> GetMostBookedServices()
        {
            return _manager.GetMostBookedServices();
        }
        public Task<(ServiceOrder, RequestStatus)> CreateServiceOrder(CreateOrder objOrder)
        {
            return _manager.CreateServiceOrder(objOrder);
        }
        public Task<(List<OrdersList>, RequestStatus)> GetServiceRequestList(ServiceRequest objserviceReq)
        {
            return _manager.GetServiceRequestList(objserviceReq);
        }
        public Task<RequestStatus> RescheduleService(ReschuduleService ReSechService)
        {
            return _manager.RescheduleService(ReSechService);
        }
        public Task<(ServiceRequestData, RequestStatus)> GetServiceRequestDetails(OrderDetailRequest objservice)
        {
            return _manager.GetServiceRequestDetails(objservice);
        }
    }
}
