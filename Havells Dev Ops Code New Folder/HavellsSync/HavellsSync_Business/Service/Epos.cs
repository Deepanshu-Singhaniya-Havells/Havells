using HavellsSync_Data.IManager;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Epos;

namespace HavellsSync_Business.IService
{
    public class Epos: IEpos
    {
        private readonly IEposManager _ePOSmanager;
        public Epos(IEposManager POSmanager)
        {
            Check.Argument.IsNotNull(nameof(POSmanager), POSmanager);
            _ePOSmanager = POSmanager;
        }

        public async Task<EposUserinfoDetails> GetConsumerInfo(string MobileNumber)
        {
            return await _ePOSmanager.GetConsumerInfo(MobileNumber);
        }

        public async Task<EposJobStatus> GetServiceCallStatus(string JobId)
        {
            return await _ePOSmanager.GetServiceCallStatus(JobId);
        }

        public async Task<ServiceCallRequestData> SyncEPOSSalesData(ServiceCallRequestData param)
        {
            return await _ePOSmanager.SyncEPOSSalesData(param);
        }

    }
}
