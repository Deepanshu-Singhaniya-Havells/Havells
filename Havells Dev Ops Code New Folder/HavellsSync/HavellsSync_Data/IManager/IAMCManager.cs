using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Product;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Data.IManager
{
    public interface IAMCManager
    {
        Task<(WarrantyContentRes, RequestStatus)> GetWarrantyContent(string SourceType, string SessionId);
        Task<(AMCPlanRes, RequestStatus)> GetAMCPlan(AMCPlanParam AMCPlanParam, string SessionId);
        Task<(PaymentStatusRes, RequestStatus)> GetStatus(PaymentStatusParam PaymentStatusParam, string SessionId);
        Task<(AMCProductDeatilsList, RequestStatus)> GetRegisteredProductList(string SourceType, string SessionId);
        Task<(AMCOrdersListRes, RequestStatus)> GetAMCOrders(AMCOrdersParam AMCOrdersParam, string SessionId);
        Task<(InitiatePaymentRes, RequestStatus)> InitiatePayment(InitiatePaymentParam InitiatePaymentParam, string SessionId);
        Task<(List<TranscationHistory>, RequestStatus)> GetTransactionDetails(AMCOrdersParam AMCOrdersParam, string MobileNumber); 

    }
}
