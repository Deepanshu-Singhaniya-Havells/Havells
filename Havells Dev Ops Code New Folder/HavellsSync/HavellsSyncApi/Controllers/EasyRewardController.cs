using HavellsSync_Business.IService;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.EasyReward;
using HavellsSync_ModelData.ICommon;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.Eventing.Reader;
using System.Diagnostics.Metrics;
using System.Net;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security.AccessControl;
using System.ServiceModel.Channels;

namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class EasyRewardController : ControllerBase
    {

        private readonly ICustomLog _logger;
        private IEasyReward _IeReward;
        private static readonly string[] SourceType = new string[] { "1", "3", "5", "6", "7", "9", "12" };
        public EasyRewardController(ICustomLog logger, IEasyReward IeReward, IAES256 AES256)
        {
            _logger = logger;
            Check.Argument.IsNotNull(nameof(IeReward), IeReward);
            this._IeReward = IeReward;
        }

        [HttpPost]
        public async Task<IActionResult> GetUserInfo([FromBody] Loyaltyparam<LoyaltySourceType> param, [FromHeader] String LoginUserId)
        {
            try
            {
                _logger.LogToFile(new Exception(string.Format("GetUserInfo|Controller|1|{0}|{1}", LoginUserId, JsonConvert.SerializeObject(param))));
                UserinfoDetails OBJUserinfoDetails = new UserinfoDetails();
                if (LoginUserId.Length != 10)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
                }
                if (param.data == null)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, CommonMessage.BadRequestMsg } });
                }
                if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }

                OBJUserinfoDetails = await _IeReward.GetUserInfo(LoginUserId);
                if (OBJUserinfoDetails == null)
                {
                    return StatusCode(StatusCodes.Status403Forbidden, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.CustomerNotExitMsg, Guid = "00000000-0000-0000-0000-000000000000" } });
                }
                _logger.LogToFile(new Exception(string.Format("GetUserInfo|Controller|10|{0}|Return From GetUserInfo", LoginUserId)));

                return StatusCode(StatusCodes.Status200OK, new { data = OBJUserinfoDetails });
            }
            catch (Exception ex)
            {
                _logger.LogToFile(new Exception(string.Format("GetUserInfo|Controller|11|{0}|Controller Exception: {1}", LoginUserId, JsonConvert.SerializeObject(ex))));
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {
                    data = new
                    {
                        Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }

        [HttpPost]
        public async Task<IActionResult> UpdateLoyaltyStatus([FromBody] Loyaltyparam<LoyaltySourceType> param, [FromHeader] String LoginUserId)
        {
            try
            {
                _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Controller|1|{0}|{1}", LoginUserId, JsonConvert.SerializeObject(param))));
                EasyRewardResponse Response = new EasyRewardResponse();
                if (LoginUserId.Length != 10)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
                }
                if (param.data == null)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }
                else if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }

                Response = await _IeReward.UpdateLoyaltyStatus(LoginUserId, param.data.SourceType);
                if (string.IsNullOrWhiteSpace(Response.Response))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.CustomerNotExitMsg } });
                }
                else
                {
                    _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Controller|7|{0}|Return From UpdateLoyaltyStatus", LoginUserId)));
                    return StatusCode(StatusCodes.Status200OK, new { data = new { StatusCode = StatusCodes.Status200OK, Message = CommonMessage.SuccessMsg } });
                }
            }
            catch (Exception ex)
            {
                _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Controller|8|{0}|Controller Exception: {1}", LoginUserId, JsonConvert.SerializeObject(ex))));
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new { Message = "Error" });
            }
        }
    }
}
