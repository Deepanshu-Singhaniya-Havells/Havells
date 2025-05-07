using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace D365WebJobs.WhatsAppCampaigns
{
    class WhatsAppCampaignDTOs
    {

    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
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
}
