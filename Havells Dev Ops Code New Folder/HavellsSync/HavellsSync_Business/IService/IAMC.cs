using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Product;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Busines.IService
{
    public interface IAMC
    {
        Task<(WarrantyContentRes, RequestStatus)> GetWarrantyContent(string SourceType, string MobileNumber);
        Task<(AMCPlanRes, RequestStatus)> GetAMCPlan(AMCPlanParam AMCPlanParam, string MobileNumber);
        Task<(PaymentStatusRes, RequestStatus)> GetStatus(PaymentStatusParam PaymentStatusParam, string MobileNumber);
        Task<(AMCProductDeatilsList, RequestStatus)> GetRegisteredProductList(string SourceType, string MobileNumber);
        Task<(AMCOrdersListRes, RequestStatus)> GetAMCOrders(AMCOrdersParam AMCOrdersParam, string MobileNumber);
        Task<(InitiatePaymentRes, RequestStatus)> InitiatePayment(InitiatePaymentParam IPParam, string MobileNumber);
        Task<(List<TranscationHistory>, RequestStatus)> GetTransactionDetails(AMCOrdersParam SourceType, string MobileNumber); 
    }
}
