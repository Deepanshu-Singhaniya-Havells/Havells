using HavellsSync_Business.IService;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ICommon;
using HavellsSync_ModelData.ServiceAlaCarte;
using Microsoft.AspNetCore.Mvc;
using System.Linq;
using System.Net;


namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class ServiceAlaCarteController : ControllerBase
    {
        private readonly ICustomLog _logger;
        private IServiceAlaCarte _ServiceAlaCarte;
        private static readonly string[] SourceType = new string[] { "6" };
        private static readonly string[] PaymentType = new string[] { "1", "2" };
        private static readonly string[] PreferedTime = new string[] { "1", "2", "3" };
        public ServiceAlaCarteController(ICustomLog logger, IServiceAlaCarte AlaCarte, IAES256 AES256)
        {
            _logger = logger;

            Check.Argument.IsNotNull(nameof(AlaCarte), AlaCarte);
            this._ServiceAlaCarte = AlaCarte;
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<ServiceAlaCartePlanInfo>>> GetServiceProductCategory([FromBody] param<SourceTypes> param)
        {
            try
            {
                string LoginUserId = HttpContext.Request.Headers["LoginUserId"];
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
                var objresult = await _ServiceAlaCarte.GetServiceProductCategory();
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { ProductCatagories = objresult.Item1, StatusCode = StatusCodes.Status200OK } });
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
        [HttpPost]
        public async Task<ActionResult<IEnumerable<ServiceAlaCartePlanInfo>>> GetServiceAlaCarteList([FromBody] param<ServiceListParam> param)
        {
            try
            {
                ServiceAlaCartePlanInfo OBJUserinfoDetails = new ServiceAlaCartePlanInfo();
                string LoginUserId = HttpContext.Request.Headers["LoginUserId"];
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
                if ((!Validate.IsvalidGuid(param.data.ProuctCategoryId)))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidCategory } });
                }
                if ((!Validate.IsvalidGuid(param.data.ProductSubCategoryId)))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidSubCategory } });
                }
                var objresult = await _ServiceAlaCarte.GetServiceAlaCarteList(param.data.ProductSubCategoryId);

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
        public async Task<ActionResult<IEnumerable<ServiceOrder>>> PaymentRetry([FromBody] param<PaymentRetryParam> param)
        {
            try
            {
                ServiceOrder OBJUserinfoDetails = new ServiceOrder();
                string LoginUserId = HttpContext.Request.Headers["LoginUserId"];
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
                if (!Validate.IsvalidGuid(param.data.OrderId))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidOrderMsg } });
                }
                if (!PaymentType.Contains(param.data.PaymentType.ToString()))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.PaymentType } });
                }
                var objresult = await _ServiceAlaCarte.PaymentRetry(param.data, LoginUserId);

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
        public async Task<ActionResult<IEnumerable<BookedService>>> GetMostBookedServices([FromBody] param<SourceTypes> param)
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
                if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }
                var objresult = await _ServiceAlaCarte.GetMostBookedServices();

                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { BookedServiceList = objresult.Item1, StatusCode = StatusCodes.Status200OK } });
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
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<ServiceOrder>>> CreateServiceOrder([FromBody] param<CreateOrder> param)
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
                if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }
                if (!PaymentType.Contains(param.data.PaymentType.ToString()))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid payment type" } });
                }
                if (!Validate.IsvalidGuid(param.data.CustomerId))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidCustomerGuid } });

                if (!Validate.IsvalidGuid(param.data.AddressId))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.Invalidaddress } });

                if (!Validate.IsvalidDate(param.data.PreferredDate))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.PreferredDate } });

                if (!PreferedTime.Contains(param.data.PreferredDateTime))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.PreferredDaytime } });

                if (!Validate.IsNumericGreaterThanZero(param.data.OrderValue))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.OrderValue } });
                if (!Validate.IsNumeric(param.data.DiscountAmount))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidDiscount } });

                if (!Validate.IsNumericGreaterThanZero(param.data.ReceiptAmount))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.ReceiptAmount } });
                if (Convert.ToDouble(param.data.ReceiptAmount) < Convert.ToDouble(param.data.DiscountAmount))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidDiscountMsg } });
                if (Convert.ToDouble(param.data.OrderValue) < Convert.ToDouble(param.data.ReceiptAmount))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidReceiptorOrderAmount } });

                if (param.data.ServiceList == null)
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.ServiceList } });

                var objresult = await _ServiceAlaCarte.CreateServiceOrder(param.data);

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
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {
                    data = new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<ServiceOrder>>> GetServiceRequestList([FromBody] param<ServiceRequest> param)
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
                if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }

                if (!Validate.IsvalidGuid(param.data.CustomerId))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidCustomerGuid } });

                var objresult = await _ServiceAlaCarte.GetServiceRequestList(param.data);

                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { OrderList = objresult.Item1, StatusCode = StatusCodes.Status200OK } });
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
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<ServiceOrder>>> RescheduleService([FromBody] param<ReschuduleService> param)
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
                if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }
                if (!Validate.IsvalidDate(param.data.PreferredDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.PreferredDate } });
                }
                if (!PreferedTime.Contains(param.data.PreferredDateTime))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.PreferredDaytime } });
                }
                if (!Validate.IsvalidGuid(param.data.orderId))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidCustomerGuid } });
                }

                var objresult = await _ServiceAlaCarte.RescheduleService(param.data);

                if (objresult.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = objresult });
                }
                else
                {
                    var RequestStatus = new
                    {
                        StatusCode = objresult.StatusCode,
                        Message = objresult.Message
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
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<ServiceOrder>>> GetServiceRequestDetails([FromBody] param<OrderDetailRequest> param)
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
                if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }

                if (!Validate.IsvalidGuid(param.data.OrderGuid))
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.InvalidCustomerGuid } });

                var objresult = await _ServiceAlaCarte.GetServiceRequestDetails(param.data);

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
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {
                    data = new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                    }
                });
            }
        }
    }
}
