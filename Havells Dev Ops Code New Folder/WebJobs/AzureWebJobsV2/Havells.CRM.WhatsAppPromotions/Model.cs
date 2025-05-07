using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.WhatsAppPromotions
{
    public class Model
    {
    }
    public enum ModeOfCommunication
    {
        SMS = 2,
        Whatsapp = 1
    }
    public static class ProductCategoryGUID
    {
        public static readonly string AirConditioner = "d51edd9d-16fa-e811-a94c-000d3af0694e";
        public static readonly string WaterPurifier = "72981d83-16fa-e811-a94c-000d3af0694e";
    }
    public static class WhatsAppCampaignTemplate
    {
        public static readonly string AC_AMC_14D = "14D_AC_AMC";
        public static readonly string WP_AMC_14D = "14D_WP_AMC";
        public static readonly string AC_AMC_30D = "30D_AC_AMC";
        public static readonly string WP_AMC_30D = "30D_WP_AMC";
        public static readonly string WP_AMC_7D = "7D_WP_AMC";
        public static readonly string AC_AMC_7D = "7D_AC_AMC";
    }
    public static class CallSubTypeGUID
    {
        public static readonly string BreakDown = "6560565A-3C0B-E911-A94E-000D3AF06CD4";
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    class WhatsAppCampaignDTOs
    {

    }
    public class logs
    {
        public string MobileNumber { get; set; }
        public string SerialNumber { get; set; }
        public string JobId { get; set; }
        public string ModeOfComunication { get; set; }
        public string Template { get; set; }
        public string URL { get; set; }
        public string Payload { get; set; }
        public string Response { get; set; }
    }

    public class UserDetails
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public Traits traits { get; set; }
        public List<object> tags { get; set; }
    }
    public class CampaignDetails
    {
        public string phoneNumber { get; set; }
        public string countryCode { get; set; }
        public string @event { get; set; }
        public Traits traits { get; set; }
    }
    public class Traits
    {
        public string name { get; set; }
        public string expire_on { get; set; }
        public string prd_cat { get; set; }
        public string product_model { get; set; }
        public string product_serial_number { get; set; }
        public string registration_date { get; set; }
    }
    public class D365Campaign
    {
        public string CampaignRunOn { get; set; }
        public string CampaignName { get; set; }
        public EntityReference Consumer { get; set; }
        public EntityReference Job_ID { get; set; }
        public string Mobile_Number { get; set; }
        public OptionSetValue Communication_Mode { get; set; }
        public string Customer_Name { get; set; }
        public EntityReference Product_Cat { get; set; }
        public EntityReference Model { get; set; }
        public string Message { get; set; }
        public string Template_ID { get; set; }
        public EntityReference Serial_Number { get; set; }
        public DateTime Registration_Date { get; set; }
    }

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class UserResponse
    {
        public bool result { get; set; }
        public string message { get; set; }
    }


}