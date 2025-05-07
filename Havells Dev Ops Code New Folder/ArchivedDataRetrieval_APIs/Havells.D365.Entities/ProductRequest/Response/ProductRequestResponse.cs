using Havells.D365.Entities.ProductRequest.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.ProductRequest.Response
{
    public class ProductRequestResponse
    {
        public List<dtoProductRequest> ProductRequest { get; set; }
        public bool Success { get; set; }
        public string Error { get; set; }

    }
}
