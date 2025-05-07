using Havells.D365.Entities.POStatusTracker.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.POStatusTracker.Response
{
    public class POStatusTrackerResponse
    {
        public List<dtoPOStatusTracker> POStatusTracker { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }

    }
}
