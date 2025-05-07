using Havells.D365.Entities.ActivityDetails.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.ActivityDetails.Response
{
   public class ActivityDetailsResponse
    {
        public List<dtoActivityDetails> ActivityDetails { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }

    }
}
