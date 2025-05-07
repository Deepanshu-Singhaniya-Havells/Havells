using Havells.D365.Entities.WorkOrders.Entity;
using System.Collections.Generic;

namespace Havells.D365.Entities.WorkOrders.Response
{
    public class WorkOrderResponse
    {
        public List<dtoWorkOrders> workOrders { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }
    }


    
}
