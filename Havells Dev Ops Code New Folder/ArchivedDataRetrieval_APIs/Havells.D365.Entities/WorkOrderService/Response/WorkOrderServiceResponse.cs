using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Havells.D365.Entities.WorkOrderService.Entity;
namespace Havells.D365.Entities.WorkOrderService.Response
{
    public class WorkOrderServiceResponse
    {
        public List<dtoWorkOrderService> OrderService { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }
    }
}
