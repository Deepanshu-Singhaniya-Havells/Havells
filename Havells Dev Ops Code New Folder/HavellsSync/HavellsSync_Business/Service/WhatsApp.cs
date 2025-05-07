using HavellsSync_Business.IService;
using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Business.Service
{
    public class WhatsApp : IWhatsApp
    {

        private readonly IWhatsAppManager _WhatsApp;
        public WhatsApp(IWhatsAppManager whatsApp)
        {
            Check.Argument.IsNotNull(nameof(whatsApp), whatsApp);
            this._WhatsApp = whatsApp;
        }

        public async Task<(WhatsappConnectRes, RequestStatus)> WhatsappConnect(WhatsAppModel whatsAppData)
        {
            return await _WhatsApp.WhatsappConnect(whatsAppData);
        }
    }
}
