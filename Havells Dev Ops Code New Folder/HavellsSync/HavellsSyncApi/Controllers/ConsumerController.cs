using HavellsSync_Business.IService;
using HavellsSync_Data;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Consumer;
using HavellsSync_ModelData.ICommon;
using Microsoft.AspNetCore.Mvc;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;

namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class ConsumerController : ControllerBase
    {
        private readonly ICustomLog _logger;
        private IConsumer _IeReward;
        private static readonly string[] SourceType = new string[] { "1", "3", "5", "6", "7", "12" };
        private static readonly string[] call_type = new string[] { "INSTALLATION", "BREAKDOWN" };
        public ConsumerController(ICustomLog logger, IConsumer IeReward, IAES256 AES256)
        {
            _logger = logger;

            Check.Argument.IsNotNull(nameof(IeReward), IeReward);
            this._IeReward = IeReward;
        }

        [HttpPost]
        public async Task<IActionResult> ConsumersAppRating([FromBody] Consumerparam<ConsumerSourceType> param, [FromHeader, Required, RegularExpression(@"^[0-9]{10}$", ErrorMessage = "Must be greater then zero.")] String LoginUserId)
        {
            try
            {
                ConsumerResponse Response = new ConsumerResponse();
                bool isNumber = false;
                if (param.data == null)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }
                else if (!SourceType.Contains(param.data.SourceType))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.SourcetypeMsg } });
                }
                else if (string.IsNullOrEmpty(param.data.Rating))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.RatingMsg } });
                }
                else if (param.data != null)
                {
                    isNumber = Regex.IsMatch(param.data.Rating, @"\b([1-9]|10)\b");
                    if (isNumber == false)
                    {
                        return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.RatingMsg } });
                    }
                }


                Response = await _IeReward.ConsumersAppRating(LoginUserId, param.data.SourceType, param.data.Rating, param.data.Review);
                if (string.IsNullOrWhiteSpace(Response.Response))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.CustomerNotExitMsg } });
                }
                else
                {
                    return StatusCode(StatusCodes.Status200OK, new { data = new { StatusCode = StatusCodes.Status200OK, Message = CommonMessage.SuccessMsg } });
                }
            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new { Message = "Error" });
            }



        }

        [HttpPost]
        public async Task<IActionResult> InvoiceDetails([FromBody] Invoiceparamdata<Invoiceparam> param)
        {
            try
            {
                InvoiceResponse Response = new InvoiceResponse();
                if (param.Data == null)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { Data = new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.BadRequestMsg } });
                }
                else if (string.IsNullOrEmpty(param.Data.FromDate) && string.IsNullOrEmpty(param.Data.ToDate) && string.IsNullOrEmpty(param.Data.OrderNumber))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { Data = new { Status = false, Message = "Invalid Input!!" } });
                }
                else if (param.Data.FromDate != "" && string.IsNullOrEmpty(param.Data.ToDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { Data = new { Status = false, Message = "Please Fill Todate!!" } });
                }
                else if (param.Data.ToDate != "" && string.IsNullOrEmpty(param.Data.FromDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { Data = new { Status = false, Message = "Please Fill Fromdate!!" } });
                }

                Response = await _IeReward.InvoiceDetails(param.Data.FromDate, param.Data.ToDate, param.Data.OrderNumber);
                if (string.IsNullOrWhiteSpace(Response.Response))
                {

                    return StatusCode(StatusCodes.Status400BadRequest, new { Data = new { Status = false, Message = "Data Not Found!!" } });
                }
                else
                {
                    return StatusCode(StatusCodes.Status200OK, new { Data = Response.Data, Status = true, Message = "Success" });

                }
            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new { Message = "Error" });
            }



        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<RequestStatus>>> PriceList([FromBody] Invoiceparamdata<List<PriceListParam>> param)
        {
            try
            {
                var objdata = param.Data.Where(s => s.KONWA == "" || s.DATBI == "" || s.KSCHL == "" || s.MATNR == "" || s.DATAB == "" || s.KBETR.ToString() == "").ToList();
                if (objdata.Count > 0)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { data = new { StatusCode = StatusCodes.Status400BadRequest, CommonMessage.Requiremessage } });
                }

                var objresult = await _IeReward.PriceList(param.Data);
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
        public async Task<ActionResult<IEnumerable<WorkOrderResponse>>> GetWorkOrdersStatus([FromBody] WorkOrderRequest objreq)
        {
            try
            {
                Regex regexDate = new Regex(@"^\d{4}\-(0[1-9]|1[012])\-(0[1-9]|[12][0-9]|3[01])$");

                if (string.IsNullOrWhiteSpace(objreq.DealerCode))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.DealerCodeRequired });
                }
                if (objreq.DealerCode.Length < 5 || objreq.DealerCode.Length > 10)
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.DealerCodeLength });
                }
                if (string.IsNullOrWhiteSpace(objreq.FromDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.FromdateRequired });
                }
                if (string.IsNullOrWhiteSpace(objreq.ToDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.TodateRequired });
                }
                if (!regexDate.IsMatch(objreq.FromDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.MandatotyDateMsg });
                }
                if (!regexDate.IsMatch(objreq.ToDate))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.MandatotyDateMsg });
                }
                string DateValidationMessage = CommonMethods.ValidateTwoDates(Convert.ToDateTime(objreq.ToDate), Convert.ToDateTime(objreq.FromDate));
                if (!string.IsNullOrWhiteSpace(DateValidationMessage))
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = DateValidationMessage });
                var objresult = await _IeReward.GetWorkOrdersStatus(objreq);
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, new { objresult.Item2.StatusCode, objresult.Item2.Message, objresult.Item1.workOrderInfos });
                }
                else
                {
                    var RequestStatus = new
                    {
                        StatusCode = objresult.Item2.StatusCode,
                        Message = objresult.Item2.Message
                    };
                    return StatusCode(StatusCodes.Status400BadRequest, RequestStatus);
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
        public async Task<ActionResult<IEnumerable<JobStatusDTO>>> GetJobstatus([FromBody] JobStatusDTO objreq)
        {
            JobStatusDTO objresult = new JobStatusDTO();
            try
            {
                if (string.IsNullOrWhiteSpace(objreq.job_id))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = CommonMessage.JobRequired });
                }

                objresult = await _IeReward.GetJobstatus(objreq);
                if (objresult.status_description == "OK.")
                {
                    return StatusCode(StatusCodes.Status200OK, objresult);
                }
                else
                {
                    return StatusCode(StatusCodes.Status200OK, objresult);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status200OK, objresult);
            }
        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<JobRequestDTO>>> CreateServiceCallRequest([FromBody] JobRequestDTO _jobRequest)
        {
            JobRequestDTO objresult = new JobRequestDTO();
            try
            {
                if (string.IsNullOrWhiteSpace(_jobRequest.customer_mobileno))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Customer Mobile Number is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                else if (!Regex.IsMatch(_jobRequest.customer_mobileno, @"^[6-9]\d{9}$"))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Incorrect Mobile Number";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                if (string.IsNullOrWhiteSpace(_jobRequest.customer_firstname))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Customer First Name is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                if (string.IsNullOrWhiteSpace(_jobRequest.address_line1))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Address Line 1 is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                if (string.IsNullOrWhiteSpace(_jobRequest.pincode))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Pincode is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                else if (!Regex.IsMatch(_jobRequest.pincode, @"^\d{6}$"))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Invalid Pincode.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }

                if (string.IsNullOrWhiteSpace(_jobRequest.call_type))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Call type is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                else
                {
                    if (!call_type.Contains(_jobRequest.call_type.ToUpper()))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Call type is not valid.";
                        return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                    }
                }
                if (string.IsNullOrWhiteSpace(_jobRequest.product_subcategory))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Product Subcategory is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                if (string.IsNullOrWhiteSpace(_jobRequest.caller_type))
                {
                    _jobRequest.status_description = "Caller Type is required.";
                    _jobRequest.status_code = "204";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                else
                {
                    if (_jobRequest.caller_type.ToUpper() != "DEALER")
                    {
                        _jobRequest.status_description = "Invalid Caller Type.";
                        _jobRequest.status_code = "204";
                        return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                    }
                }
                if (string.IsNullOrWhiteSpace(_jobRequest.dealer_code))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Dealer Code is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                DateTime preferreddate;
                if (string.IsNullOrWhiteSpace(_jobRequest.expected_delivery_date))
                {
                    _jobRequest.status_code = "204";
                    _jobRequest.status_description = "Expected Delivery Date is required.";
                    return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                }
                else
                {
                    if (!DateTime.TryParseExact(_jobRequest.expected_delivery_date, "MM-dd-yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out preferreddate))
                    {
                        _jobRequest.status_code = "204";
                        _jobRequest.status_description = "Expected Delivery Date is not in the correct format (MM-dd-yyyy)";
                        return StatusCode(StatusCodes.Status400BadRequest, _jobRequest);
                    }
                }
                objresult = await _IeReward.CreateServiceCallRequest(_jobRequest);
                if (objresult.status_code == "200")
                {
                    return StatusCode(StatusCodes.Status200OK, objresult);
                }
                else
                {
                    return StatusCode(StatusCodes.Status400BadRequest, objresult);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status200OK, objresult);
            }
        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<List<JobOutput>>>> GetJobs([FromBody] Job job)
        {
            var objresult = await _IeReward.GetJobs(job);
            return StatusCode(StatusCodes.Status200OK, objresult);

        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<List<IoTServiceCallResult>>>> IoTGetServiceCalls([FromBody] IotServiceCall job)
        {
            var objresult = await _IeReward.IoTGetServiceCalls(job);
            return StatusCode(StatusCodes.Status200OK, objresult);

        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<List<IoTRegisteredProducts>>>> IoTRegisteredProducts([FromBody] IoTRegisteredProducts registeredProduct)
        {
            var objresult = await _IeReward.IoTRegisteredProducts(registeredProduct);
            return StatusCode(StatusCodes.Status200OK, objresult);

        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<ReturnResult>>> IoTRegisterConsumer([FromBody] IoT_RegisterConsumer consumer)
        {
            var objresult = await _IeReward.IoTRegisterConsumer(consumer);
            return StatusCode(StatusCodes.Status200OK, objresult);

        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<List<IoTNatureofComplaint>>>> IoTNatureOfComplaintByProdSubcategory([FromBody] IoTNatureofComplaint natureOfComplaint)
        {
            var objresult = await _IeReward.IoTNatureOfComplaintByProdSubcategory(natureOfComplaint);
            return StatusCode(StatusCodes.Status200OK, objresult);

        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<List<NatureOfComplaint>>>> NatureOfComplaint([FromBody] NatureOfComplaint natureOfComplaint)
        {
            var objresult = await _IeReward.NatureOfComplaint(natureOfComplaint);
            return StatusCode(StatusCodes.Status200OK, objresult);

        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<List<NatureOfComplaint>>>> AllNatureOfComplaints()
        {
            var objresult = await _IeReward.AllNatureOfComplaints();
            return StatusCode(StatusCodes.Status200OK, objresult);

        }


        [HttpPost]
        public async Task<ActionResult<IEnumerable<(OCLDetailsResponse, RequestStatus)>>> GetOCLDetails([FromBody] OCLDetailsParam objreq)
        {

            try
            {
                if (string.IsNullOrWhiteSpace(objreq.OrderNumber))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "Order Number is required" });
                }

                var objResponse = await _IeReward.GetOCLDetails(objreq);
                if (objResponse.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, objResponse.Item1);

                }
                else
                {
                    var RequestStatus = new

                    {

                        StatusCode = objResponse.Item2.StatusCode,

                        Message = objResponse.Item2.Message

                    };

                    return StatusCode(StatusCodes.Status400BadRequest, RequestStatus);
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
