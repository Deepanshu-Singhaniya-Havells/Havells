using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.WorkorderProduct.Request
{
    public class WorkorderProductRequest
    {
        public string WorkOrderIncidentId { get; set; }
        public string WorkOrderProductId { get; set; }
        public string WorkOrderProduct { get; set; }
        public int PageNo { get; set; }
        public int Size { get; set; }
    }
}
