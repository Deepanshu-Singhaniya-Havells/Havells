using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Havells.D365.Entities.WorkorderProduct.Entity;
namespace Havells.D365.Entities.WorkorderProduct.Response
{
    public class WorkOrderProductResponse
    {
        public List<dtoWorkorderProduct> incident { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }
    }
}
