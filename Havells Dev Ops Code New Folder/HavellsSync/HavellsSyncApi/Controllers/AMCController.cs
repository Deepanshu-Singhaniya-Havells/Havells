using HavellsSync_Busines.IService;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ICommon;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.SignalR.Protocol;
using Newtonsoft.Json;
using System.Net;
using System.ServiceModel.Channels;

[ApiController]
[Route("api/[controller]/[action]")]
public class AMCController : ControllerBase
{
    private readonly ICustomLog _logger;
    private readonly IAMC _amcService;
    private readonly IAES256 _AES256;
    public AMCController(ICustomLog logger, IAMC amcService, IAES256 AES256)
    {
        _logger = logger;
        _AES256 = AES256;
        Check.Argument.IsNotNull(nameof(amcService), amcService);
        this._amcService = amcService;
    }

    #region MobileNumber Not Used

    [HttpPost(Name = "GetWarrantyContent")]
    public async Task<ActionResult<IEnumerable<WarrantyContentRes>>> GetWarrantyContent([FromBody] CommonReq<SourceTypeParam> param)
    {
        try
        {

            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];

            if (param.data.SourceType != "6" || string.IsNullOrEmpty(MobileNumber) || !ModelState.IsValid)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
            }
            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }

            var objresult = await _amcService.GetWarrantyContent(param.data.SourceType, "");
            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = RequestStatus });
            }
        }
        catch (Exception ex)
        {
            _logger.LogToFile(ex);
            return StatusCode((int)HttpStatusCode.ExpectationFailed, new
            {
                data = new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                }
            });
        }
    }

    [HttpPost(Name = "GetAMCPlan")]
    public async Task<ActionResult<IEnumerable<AMCPlanRes>>> GetAMCPlan([FromBody] CommonReq<AMCPlanParam> param)
    {
        try
        {
            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];

            if (param.data.SourceType != "6" || string.IsNullOrWhiteSpace(param.data.SourceType))
            {
                string msg = CommonMessage.SourcetypeMsg;
                if (!string.IsNullOrEmpty(param.data.SourceType))
                    msg = CommonMessage.InvalidSourceTypeMsg;
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = msg } });
            }

            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }
            if (string.IsNullOrEmpty(param.data.ModelNumber) || !Validate.IsValidstring(param.data.ModelNumber) || !(param.data.ModelNumber.Length < 20))
            {
                string msg = CommonMessage.ModelNumberRequired;
                if (!string.IsNullOrEmpty(param.data.ModelNumber) || !(param.data.ModelNumber.Length < 20))
                    msg = CommonMessage.InvalidModelNumber;
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = msg } });
            }
            if (param.data.AddressId == Guid.Empty.ToString() || !Validate.IsvalidGuid(param.data.AddressId))
            {
                string msg = CommonMessage.MandatotyAddressIdMsg;
                if (!string.IsNullOrWhiteSpace(param.data.AddressId))
                    msg = CommonMessage.InvalidAssestIdMsg;
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = msg } });
            }
            if (param.data.CustomerAssestId == Guid.Empty.ToString() || !Validate.IsvalidGuid(param.data.CustomerAssestId))
            {
                string msg = CommonMessage.MandatotyAssestIdMsg;
                if (!string.IsNullOrWhiteSpace(param.data.CustomerAssestId))
                    msg = CommonMessage.InvalidAssestIdMsg;
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = msg } });
            }
            var objresult = await _amcService.GetAMCPlan(param.data, "");

            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = RequestStatus });
            }
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

    [HttpPost(Name = "GetStatus")]
    public async Task<ActionResult<IEnumerable<PaymentStatusRes>>> GetStatus([FromBody] CommonReq<PaymentStatusParam> param)
    {

        try
        {
            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];

            if (param.data.SourceType != "6" || param.data.InvoiceID == new Guid() || !ModelState.IsValid)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
            }
            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }

            var objresult = await _amcService.GetStatus(param.data, "");
            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = RequestStatus });
            }
        }
        catch (Exception ex)
        {
            _logger.LogToFile(ex);
            return StatusCode((int)HttpStatusCode.ExpectationFailed, new
            {
                data = new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                }
            });
        }
    }

    [HttpPost(Name = "GetAMCOrders")]
    public async Task<ActionResult<IEnumerable<AMCOrdersListRes>>> GetAMCOrders([FromBody] CommonReq<AMCOrdersParam> param)
    {
        try
        {
            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];
            if (param.data.SourceType != "6" || string.IsNullOrEmpty(param.data.CustomerGuId) || !ModelState.IsValid)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
            }
            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }

            var objresult = await _amcService.GetAMCOrders(param.data, "");
            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = RequestStatus });
            }
        }
        catch (Exception ex)
        {
            _logger.LogToFile(ex);
            return StatusCode((int)HttpStatusCode.ExpectationFailed, new
            {
                data = new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                }
            });
        }
    }
    #endregion

    #region MobileNumber Used 
    [HttpPost(Name = "GetRegisteredProductList")]
    public async Task<ActionResult<AMCProductDeatilsList>> GetRegisteredProductList([FromBody] CommonReq<SourceTypeParam> param)
    {

        try
        {

            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];
            if (string.IsNullOrEmpty(param.data.SourceType) || string.IsNullOrEmpty(MobileNumber) || !ModelState.IsValid)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
            }
            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }

            if (param.data.SourceType != "6")
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
            }
            var objresult = await _amcService.GetRegisteredProductList(param.data.SourceType, MobileNumber);
            if (objresult.Item2.StatusCode == (int)StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = (RequestStatus) });
            }
        }
        catch (Exception ex)
        {
            _logger.LogToFile(ex);
            return StatusCode((int)HttpStatusCode.ExpectationFailed, new
            {
                data = (new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                })
            });
        }
    }

    [HttpPost(Name = "InitiatePayment")]
    public async Task<ActionResult<IEnumerable<InitiatePaymentRes>>> InitiatePayment([FromBody] CommonReq<InitiatePaymentParam> param)
    {

        try
        {
            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];
            if (!ModelState.IsValid || string.IsNullOrEmpty(MobileNumber))
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.ModelNumberRequired } });
            }
            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }

            var objresult = await _amcService.InitiatePayment(param.data, MobileNumber);
            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = RequestStatus });
            }
        }
        catch (Exception ex)
        {
            _logger.LogToFile(ex);
            return StatusCode((int)HttpStatusCode.ExpectationFailed, new
            {
                data = new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                }
            });
        }
    }

    [HttpPost(Name = "GetTransactionDetails")]
    public async Task<ActionResult<IEnumerable<List<TranscationHistory>>>> GetTransactionDetails(CommonReq<AMCOrdersParam> param)
    {
        try
        {
            string MobileNumber = HttpContext.Request.Headers["LoginUserId"];

            if (param.data.SourceType != "6" || string.IsNullOrEmpty(param.data.CustomerGuId) || !ModelState.IsValid)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
            }
            if (MobileNumber.Length != 10)
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidMobileNumber } });
            }

            var objresult = await _amcService.GetTransactionDetails(param.data, "");
            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1, StatusCode = StatusCodes.Status200OK });
            }
            else
            {
                var RequestStatus = new
                {
                    StatusCode = objresult.Item2.StatusCode,
                    Message = objresult.Item2.Message
                };
                return StatusCode(StatusCodes.Status400BadRequest, new { data = RequestStatus });
            }
        }
        catch (Exception ex)
        {
            _logger.LogToFile(ex);
            return StatusCode((int)HttpStatusCode.ExpectationFailed, new
            {
                data = new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                }
            });
        }
    }

    #endregion

}