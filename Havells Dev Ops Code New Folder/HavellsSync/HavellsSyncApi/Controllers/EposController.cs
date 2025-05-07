using HavellsSync_Business.IService;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Epos;
using HavellsSync_ModelData.ICommon;
using Microsoft.AspNetCore.Mvc;
using System.Net;
using System.Security.AccessControl;

namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class EposController : ControllerBase
    {

        private readonly ICustomLog _logger;
        private IEpos _IeReward;
        public EposController(ICustomLog logger, IEpos IeReward, IAES256 AES256)
        {
            _logger = logger;

            Check.Argument.IsNotNull(nameof(IeReward), IeReward);
            this._IeReward = IeReward;
        }

        [HttpPost]
        public async Task<IActionResult> GetConsumerInfo([FromBody] EposLoyaltyparam<ClsMobileNumber> param)
        {
            try
            {
                EposUserinfoDetails OBJUserinfoDetails = new EposUserinfoDetails();

                if (param.data == null || string.IsNullOrEmpty(param.data.MobileNumber))
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.BadRequestMsg } });
                }
                string mobileNumber = param.data.MobileNumber;
                if (mobileNumber.Length != 10)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.InvalidMobileNumber } });
                }
                OBJUserinfoDetails = await _IeReward.GetConsumerInfo(mobileNumber);
                if (OBJUserinfoDetails == null)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.CustomerNotExitMsg } });

                }
                return StatusCode(StatusCodes.Status200OK, new { data = OBJUserinfoDetails });

            }
            catch (Exception ex)
            {
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
        public async Task<IActionResult> GetServiceCallStatus([FromBody] EposLoyaltyparam<ClsJobId> param)
        {
            try
            {
                EposJobStatus JobStatus = new EposJobStatus();

                if (param.data == null || string.IsNullOrEmpty(param.data.JobId))
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.BadRequestMsg } });
                }
                string JobId = param.data.JobId;
                JobStatus = await _IeReward.GetServiceCallStatus(JobId);
                if (JobStatus == null)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.JobIdNotFound } });

                }
                return StatusCode(StatusCodes.Status200OK, new { data = JobStatus });

            }
            catch (Exception ex)
            {
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
        public async Task<ActionResult<ServiceCallRequestData>> SyncEPOSSalesData([FromBody] EposLoyaltyparam<ServiceCallRequestData> param)
        {
            try
            {
                ServiceCallRequestData JobStatus = new ServiceCallRequestData();
                if (param.data == null)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.BadRequestMsg } });
                }
                if (param.data.MobileNumber.Length != 10)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { StatusCode = false, Response = CommonMessage.InvalidMobileNumber } });
                }
                if (string.IsNullOrEmpty(param.data.Consent))
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { StatusCode = false, Response = CommonMessage.Consent } });
                }
                if (string.IsNullOrEmpty(param.data.AddressLine1))
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { StatusCode = false, Response = CommonMessage.Address } });
                }
                if (string.IsNullOrEmpty(param.data.PINCode))
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { StatusCode = false, Response = CommonMessage.PinCode } });
                }
                if (param.data.ServiceCallLineItem != null)
                {
                    if (param.data.ServiceCallLineItem.Count > 0)
                    {
                        foreach (var item in param.data.ServiceCallLineItem)
                        {
                            if (string.IsNullOrEmpty(item.SerialNumber) && string.IsNullOrEmpty(item.SKUCode))
                            {
                                return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.SerialandSkuCode } });
                            }
                            if (string.IsNullOrEmpty(item.PreferredDateofService))
                            {
                                return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Message = CommonMessage.PreferredDateofService } });
                            }
                            if (string.IsNullOrEmpty(item.PreferredTimeofService))
                            {
                                return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.PreferredTimeofService } });
                            }
                            if (string.IsNullOrEmpty(item.InstallationRequired))
                            {
                                return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.Installationrequired } });
                            }
                        }
                    }
                    else
                    {
                        return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.ServiceCallLineItemrequired } });
                    }
                }
                else
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = CommonMessage.ServiceCallLineItemrequired } });
                }
                JobStatus = await _IeReward.SyncEPOSSalesData(param.data);
                if (!JobStatus.Status)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { Status = false, Response = JobStatus.Response } });
                }
                return StatusCode(StatusCodes.Status200OK, new { data = new { Status = true, JobStatus.Response } });

            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {
                    data = new
                    {
                        Response = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }
    }
}
