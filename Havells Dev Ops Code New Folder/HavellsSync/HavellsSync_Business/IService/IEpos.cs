using HavellsSync_ModelData.Epos;

namespace HavellsSync_Business.IService
{
    public interface IEpos
    {
        Task<EposUserinfoDetails> GetConsumerInfo(string MobileNumber);

        Task<EposJobStatus> GetServiceCallStatus(string JobId);

        Task<ServiceCallRequestData> SyncEPOSSalesData(ServiceCallRequestData param);

        
    }
}
