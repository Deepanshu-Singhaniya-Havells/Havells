using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.Incident.Request
{
    public class IncidentRequest
    {
        public string JobId { get; set; }
        public string IncidentId { get; set; }
        public int PageNo { get; set; }
        public int Size { get; set; }
    }
}
