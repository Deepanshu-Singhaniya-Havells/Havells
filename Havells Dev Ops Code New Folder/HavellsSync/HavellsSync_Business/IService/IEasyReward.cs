using HavellsSync_ModelData.EasyReward;

namespace HavellsSync_Business.IService
{
    public interface IEasyReward
    {
        Task<UserinfoDetails> GetUserInfo (string MobileNumber);

        Task<EasyRewardResponse> UpdateLoyaltyStatus(string MobileNumber, string SourceType);
    }
}
