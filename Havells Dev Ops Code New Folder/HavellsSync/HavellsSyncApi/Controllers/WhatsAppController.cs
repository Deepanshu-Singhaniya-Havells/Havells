using HavellsSync_Busines.IService;
using HavellsSync_Business.IService;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ICommon;
using HavellsSync_ModelData.Product;
using Microsoft.AspNetCore.Mvc;

namespace HavellsSyncApi.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class WhatsAppController : ControllerBase
    {
        private readonly ICustomLog _logger;
        private readonly IWhatsApp _whatsApp;
        private readonly IAES256 _AES256;
        public WhatsAppController(ICustomLog logger, IWhatsApp whatsApp, IAES256 AES256)
        {
            _logger = logger;
            Check.Argument.IsNotNull(nameof(whatsApp), whatsApp);
            this._whatsApp = whatsApp;
            _AES256 = AES256;
        }

        [HttpPost(Name = "WhatsappConnect")]
        public async Task<ActionResult> WhatsappConnect(WhatsAppModel param)
        {
            var objresult = await _whatsApp.WhatsappConnect(param);

            if (objresult.Item2.StatusCode == StatusCodes.Status200OK)
            {
                return StatusCode(StatusCodes.Status200OK, new { data = objresult.Item1 });
            }
            else
            {
                return StatusCode(StatusCodes.Status400BadRequest, new { data = objresult.Item1, Error = objresult.Item2 });
            }
        }
    }
}
