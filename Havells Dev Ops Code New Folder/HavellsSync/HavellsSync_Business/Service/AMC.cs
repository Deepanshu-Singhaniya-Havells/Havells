using HavellsSync_Busines.IService;
using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Product;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Busines.Service
{
    public class AMC : IAMC
    {
        private readonly IAMCManager _amcManager;
        public AMC(IAMCManager amcManager)
        {
            Check.Argument.IsNotNull(nameof(amcManager), amcManager);

            this._amcManager = amcManager;
        }

        public async Task<(AMCPlanRes, RequestStatus)> GetAMCPlan(AMCPlanParam AMCPlanParam, string MobileNumber)
        {
            return await _amcManager.GetAMCPlan(AMCPlanParam, MobileNumber);
        }

        public async Task<(WarrantyContentRes, RequestStatus)> GetWarrantyContent(string SourceType, string MobileNumber)
        {
            return await _amcManager.GetWarrantyContent(SourceType, MobileNumber);
        }
        public async Task<(PaymentStatusRes, RequestStatus)> GetStatus(PaymentStatusParam paymentStatusParam, string MobileNumber)
        {
            return await _amcManager.GetStatus(paymentStatusParam, MobileNumber);
        }
        public async Task<(AMCProductDeatilsList, RequestStatus)> GetRegisteredProductList(string SourceType, string MobileNumber)
        {
            return await _amcManager.GetRegisteredProductList(SourceType, MobileNumber);
        }
        public async Task<(AMCOrdersListRes, RequestStatus)> GetAMCOrders(AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            return await _amcManager.GetAMCOrders(AMCOrdersParam, MobileNumber);
        }
        public async Task<(InitiatePaymentRes, RequestStatus)> InitiatePayment(InitiatePaymentParam IPParam, string MobileNumber)
        {
            return await _amcManager.InitiatePayment(IPParam, MobileNumber);
        }
        public async Task<(List<TranscationHistory>, RequestStatus)> GetTransactionDetails(AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            return await _amcManager.GetTransactionDetails(AMCOrdersParam, MobileNumber);
        }

    }
}
