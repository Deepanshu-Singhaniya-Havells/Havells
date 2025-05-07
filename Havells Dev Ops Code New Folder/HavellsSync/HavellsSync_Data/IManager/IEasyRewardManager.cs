using HavellsSync_ModelData.EasyReward;

namespace HavellsSync_Data.IManager
{
    public interface IEasyRewardManager
    {

       Task<UserinfoDetails> GetUserInfo(string MobileNumber);

        Task<EasyRewardResponse> UpdateLoyaltyStatus(string MobileNumber, string SourceType);


    }
}
