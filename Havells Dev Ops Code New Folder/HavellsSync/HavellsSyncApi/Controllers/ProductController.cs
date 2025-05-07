using Microsoft.AspNetCore.Mvc;
using HavellsSync_ModelData.Product;
using HavellsSync_ModelData.Common;
using HavellsSync_Busines.IService;
using HavellsSyncApi.Common;
using Microsoft.AspNetCore.Authorization;
using HavellsSync_Busines.Service;
using System.Net;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using Newtonsoft.Json;
using HavellsSync_ModelData.ICommon;
using System.ServiceModel.Channels;
using Microsoft.AspNetCore.Session;
using Newtonsoft.Json.Linq;

[CustomAuthorize]
[ApiController]
[Route("api/[controller]/[action]")]
public class ProductController : ControllerBase
{
    private readonly ICustomLog _logger;
    private readonly IProduct _productService;
    private readonly IAES256 _AES256;
    public ProductController(ICustomLog logger, IProduct productService, IAES256 AES256)
    {
        _logger = logger;
        Check.Argument.IsNotNull(nameof(productService), productService);
        this._productService = productService;
        _AES256 = AES256;
    }

    [HttpPost(Name = "ValidateSerialNumber")]
    public async Task<ActionResult<IEnumerable<ValidateSerialNumResponse>>> ValidateSerialNumber([FromBody] string param)
    {
        try
        {
            string SessionId = _AES256.DecryptAES256(HttpContext.Request.Headers["AccessToken"]);
            string objSerialNumber = _AES256.DecryptAES256(param);
            if (string.IsNullOrEmpty(SessionId) || string.IsNullOrEmpty(objSerialNumber))
            {
                return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid Request!!!"
                })));
            }
            SerialNumberModel? objRegProdData;
            try
            {
                objRegProdData = JsonConvert.DeserializeObject<SerialNumberModel>(objSerialNumber);
            }
            catch (Exception)
            {
                return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid Request!!!"
                })));
            }
            var objresult = await _productService.ValidateSerialNumber(objRegProdData.SerialNumber, SessionId);
            if (objresult.StatusCode == 200)
            {
                return Ok(_AES256.EncryptAES256(JsonConvert.SerializeObject(objresult)));
            }
            else if (objresult.StatusCode == 400)
            {
                return StatusCode(objresult.StatusCode, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = objresult.Message
                })));
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
                Message = "D365 Internal Server Error : " + ex.Message,
            })));
        }
    }

    [HttpPost(Name = "RegisterProduct")]
    public async Task<ActionResult<IEnumerable<RegisterProductResponse>>> RegisterProduct([FromBody] string param)
    {
        try
        {
            string SessionId = _AES256.DecryptAES256(HttpContext.Request.Headers["AccessToken"]);
            string objProductData = _AES256.DecryptAES256(param);
            if (string.IsNullOrEmpty(SessionId) || string.IsNullOrEmpty(objProductData))
            {
                return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid Request!!!"
                })));
            }
            RegisterProductModel? objRegProdData;
            try
            {
                objRegProdData = JsonConvert.DeserializeObject<RegisterProductModel>(objProductData);
            }
            catch (Exception)
            {
                return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid Request!!!"
                })));
            }
            var objresult = await _productService.RegisterProduct(objRegProdData, SessionId);
            if (objresult.StatusCode == 200)
            {
                return Ok(_AES256.EncryptAES256(JsonConvert.SerializeObject(objresult)));
            }
            else if (objresult.StatusCode == 400)
            {
                return StatusCode(objresult.StatusCode, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = objresult.Message
                })));
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
                Message = "D365 Internal Server Error : " + ex.Message,
            })));
        }
    }

    [HttpPost(Name = "GetRegisteredProducts")]
    public async Task<ActionResult<ProductDeatilsList>> GetRegisteredProducts([FromBody] string param)
    {
        try
        {
            string SessionId = _AES256.DecryptAES256(HttpContext.Request.Headers["AccessToken"]);
            string registeredProduct = _AES256.DecryptAES256(param);
            if (string.IsNullOrEmpty(SessionId) || string.IsNullOrEmpty(registeredProduct))
            {
                return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid Request!!!"
                })));
            }
            RegisterProductModel? objRegProd;
            try
            {
                objRegProd = JsonConvert.DeserializeObject<RegisterProductModel>(registeredProduct);
            }
            catch (Exception)
            {
                return StatusCode((int)HttpStatusCode.BadRequest, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid Request!!!"
                })));
            }
            var objresult = await _productService.GetRegisteredProducts(objRegProd, SessionId);
            if (objresult.StatusCode == 200)
            {
                return Ok(_AES256.EncryptAES256(JsonConvert.SerializeObject(objresult)));
            }
            else if (objresult.StatusCode == 400)
            {
                return StatusCode(objresult.StatusCode, _AES256.EncryptAES256(JsonConvert.SerializeObject(new
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = objresult.Message
                })));
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
                Message = "D365 Internal Server Error : " + ex.Message,
            })));
        }
    }
}
