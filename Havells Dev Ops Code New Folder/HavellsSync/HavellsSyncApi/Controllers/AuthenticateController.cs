using HavellsSync_Busines.IService;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ICommon;
using HavellsSync_ModelData.Product;
using HavellsSyncApi.Common;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Net;
using System.ServiceModel.Channels;

namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class AuthenticateController : ControllerBase
    {
        //private readonly ILogger<AuthenticateController> _logger;
        private readonly ICustomLog _logger;
        private readonly IAuthentication _AuthService;
        private readonly IAES256 _AES256;

        public AuthenticateController(ICustomLog logger, IAuthentication authService, IAES256 AES256)
        {
            _logger = logger;
            Check.Argument.IsNotNull(nameof(authService), authService);
            this._AuthService = authService;
            _AES256 = AES256;
        }
        [HttpPost(Name = "LoginAuthService")]
        public ActionResult<IEnumerable<AuthResponse>> LoginAuthService([FromBody] string param)
        {
            try
            {
                string objAuthjson = _AES256.DecryptAES256(param);
                if (string.IsNullOrEmpty(objAuthjson))
                {
                    return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.BadRequestMsg,
                    })));
                }
                AuthModel? objAuth;
                try
                {
                    objAuth = JsonConvert.DeserializeObject<AuthModel>(objAuthjson);
                }
                catch (Exception)
                {
                    return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.BadRequestMsg,
                    })));
                }
                var objresult = _AuthService.AuthenticateUser(objAuth);
                if (objresult.StatusCode == 200)
                {
                    return Ok(_AES256.EncryptAES256(JsonConvert.SerializeObject(objresult)));
                }
                else if (objresult.StatusCode == 204)
                {
                    return Ok(_AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        StatusCode = objresult.StatusCode,
                        Message = objresult.Message,
                        TokenExpiresAt = objresult.TokenExpiresAt
                    })));
                }
                else if (objresult.StatusCode == 400)
                {
                    return StatusCode(objresult.StatusCode, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = objresult.Message,
                    })));
                }
                else
                {
                    return StatusCode(objresult.StatusCode, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                    {
                        Message = objresult.Message,
                    })));
                }
            }
            catch (Exception ex)
            {
                _logger.LogToFile(ex);
                return StatusCode((int)HttpStatusCode.InternalServerError, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message,
                })));
            }

        }
    }
}
