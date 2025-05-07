using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Business.IService
{
    public interface IWhatsApp
    {
        Task<(WhatsappConnectRes, RequestStatus)> WhatsappConnect(WhatsAppModel AMCPlanParam);
    }
}
