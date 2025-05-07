using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Traffilog
{
    public class Datum
    {
        public string vehicle_id { get; set; }
        public string unit_id { get; set; }
        public string unit_serial { get; set; }
        public string license_nmbr { get; set; }
        public string chassis_number { get; set; }
        public DateTime last_communication_time { get; set; }
        public DateTime last_position_time { get; set; }
        public string latitude { get; set; }
        public string longtitude { get; set; }
        public string speed { get; set; }
        public string direction { get; set; }
        public string status { get; set; }
        public DateTime last_event_time { get; set; }
        public string last_event_type { get; set; }
        public string current_driver { get; set; }
        public string current_driver_number { get; set; }
        public string driver_name { get; set; }
        public string current_drive { get; set; }
        public string last_mileage { get; set; }
    }

    public class Properties
    {
        public string action_name { get; set; }
        public List<Datum> data { get; set; }
        public string action_value { get; set; }
        public string description { get; set; }
        public string session_token { get; set; }
    }

    public class Response
    {
        public Properties properties { get; set; }
    }

    public class Root
    {
        public Response response { get; set; }
    }

}
