using System.Collections.Generic;

namespace Havells.D365.Entities.Common.Response
{
    public class CommonResponse
    {
        public IEnumerable<string> Errors { get; set; }
        public bool Success { get; set; }
    }
}
