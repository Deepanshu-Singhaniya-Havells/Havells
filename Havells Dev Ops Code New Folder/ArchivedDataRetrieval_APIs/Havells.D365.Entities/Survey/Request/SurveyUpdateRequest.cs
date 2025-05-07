using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.Survey.Request
{
    public class SurveyUpdateRequest
    {
        public string JobId { get; set; }
        public string NPSValue { get; set; }
        public string DetractorsResponse { get; set; }
        public string PassivesResponse { get; set; }
        public string PromotersResponse { get; set; }
        public string Feedback { get; set; }
        public string ServiceEngineerRating { get; set; }
        public bool SubmitStatus { get; set; }
    }
}
