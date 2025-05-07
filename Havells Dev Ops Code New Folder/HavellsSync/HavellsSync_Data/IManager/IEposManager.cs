using HavellsSync_ModelData.Epos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Data.IManager
{
    public interface IEposManager
    {
        Task<EposUserinfoDetails> GetConsumerInfo(string MobileNumber);

        Task<EposJobStatus> GetServiceCallStatus(string JobId);

        Task<ServiceCallRequestData> SyncEPOSSalesData(ServiceCallRequestData param);
    }
}
