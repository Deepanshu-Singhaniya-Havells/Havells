using Havells.D365.Entities.WorkOrders.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.WorkOrders.Response
{
   public class SAWActivityResponse
    {
        public List<dtoSAWActivity> SAWActivity { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }

    }
}
