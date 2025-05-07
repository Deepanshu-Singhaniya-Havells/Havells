using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Havells.D365.Entities.Incident.Entity;

namespace Havells.D365.Entities.Incident.Response
{
   public class IncidentResponse
    {
        public List<dtoIncidents> incident { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }
    }
}
