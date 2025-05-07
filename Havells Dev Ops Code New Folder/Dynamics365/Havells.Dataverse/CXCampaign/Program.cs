using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using static CXCampaign.CXCampaignModels;

namespace CXCampaign
{
    internal class Program
    {
        static void Main(string[] args)
        {
            CampaignEngine _campaign = new CampaignEngine();
            TriggerCampaign _triggerEvent = new TriggerCampaign();

            //#region New Buyer +14 days (WA)
            //_campaign.Campaign_NewBuyer("New Buyer +14 days (AC-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.AC_AMC_14D, -14, ProductCategoryGUID.AirConditioner);
            //_campaign.Campaign_NewBuyer("New Buyer +14 days (WP-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.WP_AMC_14D, -14, ProductCategoryGUID.WaterPurifier);
            //#endregion

            //#region New Buyer +30 days (SMS)
            //_campaign.Campaign_NewBuyer("New Buyer + 30 days(AC-SMS)", ModeOfCommunication.SMS, "1107168654225056811", -30, ProductCategoryGUID.AirConditioner);
            //_campaign.Campaign_NewBuyer("New Buyer + 30 days(WP-SMS)", ModeOfCommunication.SMS, "1107169702429550488", -30, ProductCategoryGUID.WaterPurifier);
            //#endregion

            ////#region Breakdown Job Closure +7 Days (WA)
            //_campaign.Campaign_OutWarrantyBreakdownJobClosure("Breakdown Job Closure +7 Days (AC-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.AC_AMC_7D, -7, ProductCategoryGUID.AirConditioner);
            //_campaign.Campaign_OutWarrantyBreakdownJobClosure("Breakdown Job Closure +7 Days (WP-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.WP_AMC_7D, -7, ProductCategoryGUID.WaterPurifier);
            ////#endregion

            ////#region Warranty Expiry -30 Days (WA)
            //_campaign.Campaign_WarrantyExpireNear("Warranty Expiry -30 Days (AC-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.AC_AMC_30D, 30, ProductCategoryGUID.AirConditioner);
            //_campaign.Campaign_WarrantyExpireNear("Warranty Expiry -30 Days (WP-WA)", ModeOfCommunication.Whatsapp, WhatsAppCampaignTemplate.WP_AMC_30D, 30, ProductCategoryGUID.WaterPurifier);
            ////#endregion

            //#region MoEngageJobKKGClosure (API)
            ////_campaign.Campaign_MoengageKKGClosure("MoEngage Job KKG Closure (API)", ModeOfCommunication.API, null, 0, null);
            //#endregion

            #region Trigger Campaign
            _triggerEvent.ExecuteCamapigns(ModeOfCommunication.SMS);
            _triggerEvent.ExecuteCamapigns(ModeOfCommunication.Whatsapp);
            #endregion
        }
    }
}
