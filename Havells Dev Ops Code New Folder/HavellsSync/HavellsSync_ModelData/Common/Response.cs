using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.Common
{
    public class BaseResponse
    {
        public string StatusCode { get; set; }
        public string Message { get; set; }
        public string TokenExpireAt { get; set; } 
    }
}
