using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.WorkOrderService.Request
{
    public class WorkOrderServiceRequest
    {
        public string WorkOrderID { get; set; }
        public string WorkOrderServiceID { get; set; }
        public string WorkOrderIncidentID { get; set; }
    }
}
