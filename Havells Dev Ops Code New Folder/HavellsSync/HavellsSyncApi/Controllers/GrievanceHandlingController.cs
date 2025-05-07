using HavellsSync_Business.IService;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.GrievanceHandling;
using HavellsSync_ModelData.ICommon;
using HavellsSync_ModelData.ServiceAlaCarte;
using Microsoft.AspNetCore.Mvc;
using System.Linq;
using System.Net;


namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class GrievanceHandlingController : ControllerBase
    {
        private readonly ICustomLog _logger;
        private IGrievanceHandling _GrievanceHandling;
       
        public GrievanceHandlingController(ICustomLog logger, IGrievanceHandling GrievanceHandling, IAES256 AES256)
        {
            _logger = logger;

            Check.Argument.IsNotNull(nameof(GrievanceHandling), GrievanceHandling);
            this._GrievanceHandling = GrievanceHandling;
        }
        [HttpGet]
        public async Task<ActionResult<IEnumerable<Response>>> GetComplaintCategory()
        {
            try
            {
                var objresult = await _GrievanceHandling.GetComplaintCategory();
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, objresult.Item1);
                }
                else
                {
                    var RequestStatus = new
                    {
                        StatusCode = objresult.Item2.StatusCode,
                        Message = objresult.Item2.Message
                    };
                    return StatusCode(StatusCodes.Status400BadRequest,  RequestStatus);
                }
            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {                  
                 Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()                   
                });
            }
        }

        [HttpPost]
        public async Task<ActionResult<IEnumerable<OpenJobsResponse>>> GetOpenJobs([FromBody] OpenJobsRequest obj)
        {
            try
            {
                if (string.IsNullOrEmpty(obj.CustomerGuid))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "CustomerGuid required" });

                }
                var objresult = await _GrievanceHandling.GetOpenJobs(obj);
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, objresult.Item1);
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
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<OpenComplaints_Response>>> GetOpenComplaints([FromBody] OpenComplaints_Request obj)
        {
            try
            {
                if (string.IsNullOrEmpty(obj.CustomerGuid))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "CustomerGuid required" });

                }
                var objresult = await _GrievanceHandling.GetOpenComplaints(obj);
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, objresult.Item1);
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
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<Complaints_Response>>> GetComplaints([FromBody] Complaints_Request obj)
        {
            try
            {
                if (string.IsNullOrEmpty(obj.CustomerGuid))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "CustomerGuid required" });

                }
                var objresult = await _GrievanceHandling.GetComplaints(obj);
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, objresult.Item1);
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
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<RaiseReminder_Response>>> RaiseReminder(RaiseReminder_Request obj)
        {
            try
            {
                if (string.IsNullOrEmpty(obj.ComplaintGuid))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "CustomerGuid required" });

                }
                var objresult = await _GrievanceHandling.RaiseReminder(obj);
                if (objresult.Item1.Status == true)
                {
                    return StatusCode(StatusCodes.Status200OK, objresult.Item1);
                }
                else
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, objresult.Item1);
                }
            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                });
            }
        }
        [HttpPost]
        public async Task<ActionResult<IEnumerable<CreateCase_Response>>> CreateCase(CreateCase_Request obj)
        {
            try
            {
               
                if (string.IsNullOrEmpty(obj.AddressGuid))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "AddressGuid cannot be empty" });
                   
                }
                //if (string.IsNullOrEmpty(obj.ServiceRequestGuid))
                //{
                //    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "ServiceRequestGuid cannot be empty" });
                //}
                if (string.IsNullOrEmpty(obj.ComplaintCategoryId))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "ComplaintCategoryId cannot be empty" });
                }
                if (string.IsNullOrEmpty(obj.ComplaintTitle))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "ComplaintTitle cannot be empty" });
                }
                if (string.IsNullOrEmpty(obj.CustomerGuid))
                {
                    return StatusCode(StatusCodes.Status400BadRequest, new { StatusCode = StatusCodes.Status400BadRequest, Message = "CustomerGuid required" });

                }
                var objresult = await _GrievanceHandling.CreateCase(obj);
                if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
                {
                    return StatusCode(StatusCodes.Status200OK, objresult.Item1);
                }
                else
                {
                    return StatusCode(StatusCodes.Status500InternalServerError, objresult.Item2);
                }
            }
            catch (Exception ex)
            {
                return StatusCode((int)HttpStatusCode.ExpectationFailed, new
                {
                    Message = CommonMessage.InternalServerErrorMsg + " " + ex.Message.ToUpper()
                });
            }
        }

    }
}
