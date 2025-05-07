using HavellsSync_Business.IService;
using HavellsSync_Data.IManager;
using HavellsSync_Data.Manager;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.EasyReward;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Business.IService
{
    public class EasyReward : IEasyReward
    {
        private IEasyRewardManager _manager;
        public EasyReward(IEasyRewardManager ERManager)
        {
            Check.Argument.IsNotNull(nameof(ERManager), ERManager);
            _manager = ERManager;
        }
        public async Task<UserinfoDetails> GetUserInfo(string MobileNumber)
        {
            return await _manager.GetUserInfo(MobileNumber);
        }

        public async Task<EasyRewardResponse> UpdateLoyaltyStatus(string MobileNumber, string SourceType)
        {
            return await _manager.UpdateLoyaltyStatus(MobileNumber, SourceType);
        }
    }
}
